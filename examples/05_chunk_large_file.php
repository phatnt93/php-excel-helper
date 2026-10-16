<?php

declare(strict_types=1);

require dirname(__DIR__) . '/vendor/autoload.php';

use ExcelHelper\Excel;
use ExcelHelper\Reader\ChunkContext;

$tempPath = sys_get_temp_dir() . '/example_05_' . uniqid() . '.csv';

$csvData = "id,action,logged_at\n";
for ($i = 1; $i <= 250; $i++) {
    $csvData .= "LOG-{$i},user_login,2026-10-16 10:00:00\n";
}
file_put_contents($tempPath, $csvData);

$summary = Excel::chunk(
    path: $tempPath,
    options: [
        'format' => 'csv',
        'header_row' => 1,
        'columns' => [
            ['header' => 'id', 'key' => 'id'],
            ['header' => 'action', 'key' => 'action'],
            ['header' => 'logged_at', 'key' => 'time', 'type' => 'datetime'],
        ],
    ],
    size: 100,
    callback: function (array $rows, ChunkContext $ctx): void {
        echo "Processing chunk #{$ctx->chunkIndex()}: {$ctx->rowCount} rows (Rows {$ctx->startRow}..{$ctx->endRow})\n";
    },
);

echo "Finished reading in chunks!\n";
echo "Total rows: {$summary->totalRows()}, Total chunks: {$summary->totalChunks()}\n";

@unlink($tempPath);
exit(0);
