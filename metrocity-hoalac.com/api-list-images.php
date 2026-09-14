<?php

declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');

function json_response(int $statusCode, array $payload): void
{
    http_response_code($statusCode);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'GET') {
    json_response(405, [
        'ok' => false,
        'message' => 'Chỉ hỗ trợ GET.'
    ]);
}

$source = trim((string)($_GET['source'] ?? ''));
$imagesRootPath = $source === 'zones'
    ? __DIR__ . DIRECTORY_SEPARATOR . '..' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'images cac khu dat'
    : __DIR__ . DIRECTORY_SEPARATOR . 'images';
$imagesRoot = realpath($imagesRootPath);
if ($imagesRoot === false || !is_dir($imagesRoot)) {
    json_response(500, [
        'ok' => false,
        'message' => 'Không tìm thấy thư mục images.'
    ]);
}

$query = trim((string)($_GET['q'] ?? ''));
$queryLower = mb_strtolower($query, 'UTF-8');

$kind = mb_strtolower(trim((string)($_GET['kind'] ?? 'all')), 'UTF-8');
if (!in_array($kind, ['all', 'image', 'video'], true)) {
    $kind = 'all';
}

$imageExt = ['jpg', 'jpeg', 'png', 'webp', 'gif', 'svg', 'avif'];
$videoExt = ['mp4', 'mov', 'webm'];
$allowedExt = array_merge($imageExt, $videoExt);
$items = [];

$iter = new RecursiveIteratorIterator(
    new RecursiveDirectoryIterator($imagesRoot, FilesystemIterator::SKIP_DOTS)
);

foreach ($iter as $fileInfo) {
    /** @var SplFileInfo $fileInfo */
    if (!$fileInfo->isFile()) {
        continue;
    }

    $ext = mb_strtolower((string)$fileInfo->getExtension(), 'UTF-8');
    if (!in_array($ext, $allowedExt, true)) {
        continue;
    }
    if ($kind === 'image' && !in_array($ext, $imageExt, true)) {
        continue;
    }
    if ($kind === 'video' && !in_array($ext, $videoExt, true)) {
        continue;
    }

    $fullPath = $fileInfo->getPathname();
    $relPath = substr($fullPath, strlen($imagesRoot) + 1);
    if ($relPath === false) {
        continue;
    }

    $relPath = str_replace(DIRECTORY_SEPARATOR, '/', $relPath);
    $displayPath = $source === 'zones'
        ? '../images/images cac khu dat/' . $relPath
        : 'images/' . $relPath;

    if ($queryLower !== '') {
        $name = mb_strtolower($fileInfo->getBasename(), 'UTF-8');
        $pathLower = mb_strtolower($displayPath, 'UTF-8');
        if (mb_strpos($name, $queryLower) === false && mb_strpos($pathLower, $queryLower) === false) {
            continue;
        }
    }

    $normalized = str_replace(' ', '%20', $displayPath);
    $items[] = $normalized;
}

sort($items, SORT_NATURAL | SORT_FLAG_CASE);

if (count($items) > 300) {
    $items = array_slice($items, 0, 300);
}

json_response(200, [
    'ok' => true,
    'count' => count($items),
    'items' => $items
]);
