<?php
/**
 * Vehicle market price comparison app.
 * Run locally with: php -S localhost:8000
 */
declare(strict_types=1);

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

const MAX_FILE_SIZE = 10 * 1024 * 1024;

function h(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

function normalizeHeader(string $header): string
{
    $header = preg_replace('/^\\xEF\\xBB\\xBF/', '', trim($header)) ?? trim($header);
    return mb_strtolower($header, 'UTF-8');
}

function parsePrice(mixed $value): ?float
{
    if ($value === null || trim((string)$value) === '') return null;
    if (is_numeric($value)) return (float)$value;
    $clean = preg_replace('/[^0-9,.-]/u', '', (string)$value) ?? '';
    if ($clean === '' || $clean === '-') return null;
    // Indonesian currency convention: dots are thousands separators and comma is decimal.
    if (str_contains($clean, ',') && str_contains($clean, '.')) {
        $clean = strrpos($clean, ',') > strrpos($clean, '.')
            ? str_replace(',', '.', str_replace('.', '', $clean))
            : str_replace(',', '', $clean);
    } elseif (str_contains($clean, ',')) {
        $clean = str_replace(',', '.', $clean);
    } else {
        $clean = str_replace('.', '', $clean);
    }
    return is_numeric($clean) ? (float)$clean : null;
}

/** Read CSV using semicolon or comma delimiter, or XLSX/XLS using PhpSpreadsheet. */
function readRows(array $file, string $priceHeader): array
{
    if (($file['error'] ?? UPLOAD_ERR_NO_FILE) !== UPLOAD_ERR_OK) {
        throw new RuntimeException('Upload gagal. Pastikan file dipilih dan ukurannya tidak melebihi 10 MB.');
    }
    if (($file['size'] ?? 0) > MAX_FILE_SIZE) throw new RuntimeException('Ukuran file maksimal 10 MB.');
    $extension = strtolower(pathinfo((string)$file['name'], PATHINFO_EXTENSION));
    if (!in_array($extension, ['csv', 'xlsx', 'xls'], true)) {
        throw new RuntimeException('Format file harus CSV, XLSX, atau XLS.');
    }
    $records = [];
    if ($extension === 'csv') {
        $handle = fopen((string)$file['tmp_name'], 'rb');
        if (!$handle) throw new RuntimeException('File CSV tidak dapat dibaca.');
        $first = fgets($handle);
        if ($first === false) { fclose($handle); throw new RuntimeException('File CSV kosong.'); }
        $delimiter = substr_count($first, ';') >= substr_count($first, ',') ? ';' : ',';
        rewind($handle);
        $headers = fgetcsv($handle, 0, $delimiter);
        if (!$headers) { fclose($handle); throw new RuntimeException('Header CSV tidak ditemukan.'); }
        $headers = array_map(fn($v) => normalizeHeader((string)$v), $headers);
        while (($row = fgetcsv($handle, 0, $delimiter)) !== false) {
            if (count(array_filter($row, fn($v) => trim((string)$v) !== '')) === 0) continue;
            $records[] = array_combine(array_slice($headers, 0, count($row)), array_slice($row, 0, count($headers))) ?: [];
        }
        fclose($handle);
    } else {
        $book = \PhpOffice\PhpSpreadsheet\IOFactory::load((string)$file['tmp_name']);
        $sheet = $book->getActiveSheet();
        $rawHeaders = $sheet->rangeToArray('A1:' . $sheet->getHighestColumn() . '1', null, true, false)[0] ?? [];
        $headers = array_map(fn($v) => normalizeHeader((string)$v), $rawHeaders);
        for ($rowNum = 2; $rowNum <= $sheet->getHighestRow(); $rowNum++) {
            $values = $sheet->rangeToArray('A' . $rowNum . ':' . $sheet->getHighestColumn() . $rowNum, null, true, false)[0] ?? [];
            if (count(array_filter($values, fn($v) => $v !== null && trim((string)$v) !== '')) === 0) continue;
            $records[] = array_combine(array_slice($headers, 0, count($values)), array_slice($values, 0, count($headers))) ?: [];
        }
        $book->disconnectWorksheets();
    }
    if (!$records) throw new RuntimeException('File tidak memiliki data kendaraan.');
    $normalizedPrice = normalizeHeader($priceHeader);
    foreach ($records as $index => $record) {
        if (!array_key_exists('vehicle_id', $record)) throw new RuntimeException('Kolom VEHICLE_ID tidak ditemukan di ' . basename((string)$file['name']) . '.');
        if (!array_key_exists($normalizedPrice, $record)) throw new RuntimeException('Kolom ' . $priceHeader . ' tidak ditemukan di ' . basename((string)$file['name']) . '.');
        $id = trim((string)$record['vehicle_id']);
        if ($id === '') continue;
        $records[$index]['_id'] = $id;
        $records[$index]['_price'] = parsePrice($record[$normalizedPrice]);
    }
    return array_values(array_filter($records, fn($r) => isset($r['_id'])));
}

function mapById(array $rows): array
{
    $map = [];
    foreach ($rows as $row) $map[$row['_id']] = $row;
    return $map;
}

function makeComparison(array $aRows, array $bRows): array
{
    $a = mapById($aRows);
    $b = mapById($bRows);
    $ids = array_values(array_unique(array_merge(array_keys($a), array_keys($b))));
    $out = [];
    foreach ($ids as $id) {
        $left = $a[$id] ?? null;
        $right = $b[$id] ?? null;
        $model = $left['vehicle_model'] ?? $right['vehicle_model'] ?? '';
        $year = $left['tahun'] ?? $right['tahun'] ?? '';
        $priceA = $left['_price'] ?? null;
        $priceB = $right['_price'] ?? null;
        if ($left && $right && $priceA !== null && $priceB !== null) {
            $average = ($priceA + $priceB) / 2;
            $difference = $priceB - $priceA;
            $status = $difference > 0 ? 'Naik' : ($difference < 0 ? 'Turun' : 'Tidak berubah');
            $statusRecord = 'Dibandingkan';
        } elseif ($right) {
            $average = $priceB;
            $difference = null;
            $status = 'Data baru';
            $statusRecord = 'Hanya ada di File B';
        } else {
            $average = $priceA;
            $difference = null;
            $status = 'Tidak ada pembanding';
            $statusRecord = 'Hanya ada di File A';
        }
        $out[] = [$id, (string)$model, (string)$year, $priceA, $priceB, $average, $difference, $status, $statusRecord];
    }
    return $out;
}

function downloadXlsx(array $rows): never
{
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Perbandingan Harga');
    $headers = ['vehicle_id', 'vehicle_model', 'tahun', 'harga_kendaraan_sehati', 'harga_pasar_olx', 'harga_baru', 'selisih_perubahan', 'status_perubahan', 'STATUS'];
    $sheet->fromArray($headers, null, 'A1');
    if ($rows) $sheet->fromArray($rows, null, 'A2');
    $sheet->getStyle('A1:I1')->getFont()->setBold(true);
    $sheet->getStyle('A1:I1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFE2F0D9');
    $sheet->freezePane('A2');
    $sheet->setAutoFilter('A1:I' . max(1, count($rows) + 1));
    foreach (range('A', 'I') as $column) $sheet->getColumnDimension($column)->setAutoSize(true);
    foreach (['D','E','F','G'] as $column) $sheet->getStyle($column . '2:' . $column . (count($rows) + 1))->getNumberFormat()->setFormatCode('#,##0.00;[Red]-#,##0.00');
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="hasil-perbandingan-harga.xlsx"');
    header('Cache-Control: max-age=0');
    (new Xlsx($spreadsheet))->save('php://output');
    $spreadsheet->disconnectWorksheets();
    exit;
}

$errors = [];
$results = null;
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        $a = readRows($_FILES['file_a'] ?? [], 'HARGA KENDARAAN SEHATI');
        $b = readRows($_FILES['file_b'] ?? [], 'HARGA PASAR OLX');
        $results = makeComparison($a, $b);
        if (isset($_POST['download'])) downloadXlsx($results);
    } catch (Throwable $e) { $errors[] = $e->getMessage(); }
}
?>
<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Perbandingan Harga Pasar Kendaraan</title>
  <style>
    :root{font-family:Inter,system-ui,-apple-system,"Segoe UI",sans-serif;color:#172033;background:#f3f6fb}*{box-sizing:border-box}body{margin:0}.wrap{max-width:1180px;margin:42px auto;padding:0 20px}.hero{background:linear-gradient(125deg,#123b70,#2074a8);color:#fff;border-radius:18px;padding:30px 34px;box-shadow:0 10px 28px #17396024}.hero h1{margin:0 0 8px;font-size:28px}.hero p{margin:0;color:#e3efff}.card{background:#fff;border:1px solid #e5eaf2;border-radius:16px;padding:24px;margin-top:22px;box-shadow:0 5px 20px #1c31500b}.uploads{display:grid;grid-template-columns:1fr 1fr;gap:18px}.field{border:1px dashed #aebdd0;background:#f9fbfe;padding:18px;border-radius:12px}.field label{display:block;font-weight:700;margin-bottom:8px}.field small{color:#64748b;display:block;margin:6px 0 12px}.field input{width:100%}.actions{display:flex;gap:10px;margin-top:20px;flex-wrap:wrap}button{border:0;border-radius:9px;padding:12px 18px;font-weight:700;font-size:15px;cursor:pointer;background:#1769aa;color:#fff}button.secondary{background:#16804b}.alert{padding:12px 15px;border-radius:9px;margin:16px 0;background:#fff0ee;color:#a43428}.stats{color:#526176;margin:0 0 14px}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;font-size:13px}th,td{text-align:left;padding:10px 12px;border-bottom:1px solid #e7edf4;white-space:nowrap}th{background:#eff5fa;color:#26394f;position:sticky;top:0}td.num{text-align:right;font-variant-numeric:tabular-nums}.tag{display:inline-block;background:#e8f4ec;color:#17633b;padding:4px 8px;border-radius:20px}.tag.down{background:#fff0ed;color:#9e342b}.tag.new{background:#e8f0ff;color:#3155a3}.note{font-size:13px;color:#627187;line-height:1.55}.footer{color:#748197;font-size:12px;text-align:center;padding:24px}@media(max-width:700px){.wrap{margin:20px auto}.hero{padding:24px}.hero h1{font-size:23px}.uploads{grid-template-columns:1fr}}
  </style>
</head>
<body><main class="wrap">
  <header class="hero"><h1>Perbandingan Harga Pasar Kendaraan</h1><p>Unggah CSV/XLSX File A dan File B untuk membandingkan harga berdasarkan VEHICLE_ID dan mengunduh hasil Excel.</p></header>
  <section class="card">
    <?php foreach ($errors as $error): ?><div class="alert"><?= h($error) ?></div><?php endforeach; ?>
    <form method="post" enctype="multipart/form-data">
      <div class="uploads">
        <div class="field"><label for="file_a">File A — Harga Existing</label><small>Harus memiliki kolom VEHICLE_ID, VEHICLE_MODEL, TAHUN, HARGA KENDARAAN SEHATI.</small><input type="file" id="file_a" name="file_a" accept=".csv,.xlsx,.xls" required></div>
        <div class="field"><label for="file_b">File B — Harga Eksternal</label><small>Harus memiliki kolom VEHICLE_ID, VEHICLE_MODEL, TAHUN, HARGA PASAR OLX.</small><input type="file" id="file_b" name="file_b" accept=".csv,.xlsx,.xls" required></div>
      </div>
      <div class="actions"><button type="submit" name="preview" value="1">Bandingkan Harga</button><?php if ($results !== null): ?><button class="secondary" type="submit" name="download" value="1">Unduh Hasil Excel (.xlsx)</button><?php endif; ?></div>
    </form>
    <p class="note">File CSV dapat memakai pemisah titik koma (;) atau koma (,). Kendaraan yang hanya terdapat di File B akan tetap dimasukkan; harga baru mengikuti harga File B. Untuk data yang hanya ada di File A, nilai harga baru mengikuti harga File A.</p>
  </section>
  <?php if ($results !== null): ?>
  <section class="card"><p class="stats">Berhasil membandingkan <strong><?= count($results) ?></strong> kendaraan. Harga baru untuk kendaraan yang ada pada kedua file adalah rata-rata File A dan File B. Selisih = File B − File A.</p>
    <div class="tablewrap"><table><thead><tr><th>VEHICLE_ID</th><th>VEHICLE_MODEL</th><th>TAHUN</th><th>Harga Existing</th><th>Harga Eksternal</th><th>Harga Baru</th><th>Selisih</th><th>Status Perubahan</th><th>STATUS</th></tr></thead><tbody>
      <?php foreach ($results as $r): ?><tr><td><?= h((string)$r[0]) ?></td><td><?= h((string)$r[1]) ?></td><td><?= h((string)$r[2]) ?></td><td class="num"><?= $r[3] === null ? '—' : number_format((float)$r[3], 0, ',', '.') ?></td><td class="num"><?= $r[4] === null ? '—' : number_format((float)$r[4], 0, ',', '.') ?></td><td class="num"><?= $r[5] === null ? '—' : number_format((float)$r[5], 0, ',', '.') ?></td><td class="num"><?= $r[6] === null ? '—' : number_format((float)$r[6], 0, ',', '.') ?></td><td><span class="tag <?= $r[7] === 'Turun' ? 'down' : ($r[7] === 'Data baru' ? 'new' : '') ?>"><?= h((string)$r[7]) ?></span></td><td><?= h((string)$r[8]) ?></td></tr><?php endforeach; ?>
    </tbody></table></div>
  </section><?php endif; ?>
  <div class="footer">Aplikasi Perbandingan Harga Pasar</div>
</main></body></html>
