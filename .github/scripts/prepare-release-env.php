<?php

declare(strict_types=1);

/**
 * @param  array<string, string>  $variables
 */
function prepareReleaseEnvironmentFile(
    string $templatePath,
    string $outputPath,
    array $variables,
): void {
    $contents = file_get_contents($templatePath);

    if ($contents === false) {
        throw new RuntimeException(sprintf('Unable to read template file [%s].', $templatePath));
    }

    $lineEnding = detectLineEnding($contents);

    foreach ($variables as $key => $value) {
        $escapedValue = str_replace(['\\', '"'], ['\\\\', '\\"'], $value);
        $line = sprintf('%s="%s"', $key, $escapedValue);
        $pattern = '/^'.preg_quote($key, '/').'=.*$/m';

        if (preg_match($pattern, $contents) === 1) {
            $contents = (string) preg_replace($pattern, $line, $contents, 1);

            continue;
        }

        if ($contents !== '' && ! str_ends_with($contents, $lineEnding)) {
            $contents .= $lineEnding;
        }

        $contents .= $line.$lineEnding;
    }

    file_put_contents($outputPath, $contents);
}

function detectLineEnding(string $contents): string
{
    if (str_contains($contents, "\r\n")) {
        return "\r\n";
    }

    if (str_contains($contents, "\n")) {
        return "\n";
    }

    if (str_contains($contents, "\r")) {
        return "\r";
    }

    return PHP_EOL;
}

/**
 * @param  list<string>  $keys
 * @return array<string, string>
 */
function releaseEnvironmentValues(array $keys): array
{
    $values = [];

    foreach ($keys as $key) {
        $value = getenv($key);
        $values[$key] = $value === false ? '' : $value;
    }

    return $values;
}

function runPrepareReleaseEnvironmentFile(): void
{
    // 👇 СНАЧАЛА ГЕНЕРИРУЕМ КЛЮЧ (если APP_KEY не передан через env)
    $appKey = getenv('APP_KEY');
    if (empty($appKey)) {
        echo "Generating application key...\n";
        exec('php artisan key:generate --force', $output, $returnCode);
        
        if ($returnCode !== 0) {
            echo "Warning: Failed to generate APP_KEY via artisan. Using fallback.\n";
            // Fallback: генерируем ключ вручную
            $appKey = 'base64:' . base64_encode(random_bytes(32));
        } else {
            // Читаем сгенерированный ключ из .env
            if (file_exists('.env')) {
                $envContent = file_get_contents('.env');
                if (preg_match('/^APP_KEY=(.+)$/m', $envContent, $matches)) {
                    $appKey = trim($matches[1]);
                }
            }
        }
        
        // Если ключ всё ещё пуст, генерируем вручную
        if (empty($appKey)) {
            $appKey = 'base64:' . base64_encode(random_bytes(32));
        }
        
        // Устанавливаем в окружение для следующего шага
        putenv("APP_KEY=$appKey");
        $_ENV['APP_KEY'] = $appKey;
    }

    // 👇 ПОТОМ СОЗДАЁМ .env ФАЙЛ
    prepareReleaseEnvironmentFile(
        templatePath: '.env.example',
        outputPath: '.env',
        variables: releaseEnvironmentValues([
            'APP_KEY',
            'APP_ENV',
            'APP_DEBUG',
            'NATIVEPHP_APP_VERSION',
            'NATIVEPHP_APP_ID',
            'NATIVEPHP_APP_AUTHOR',
            'NATIVEPHP_APP_DESCRIPTION',
            'NATIVEPHP_APP_COPYRIGHT',
            'NATIVEPHP_APP_WEBSITE',
            'NATIVEPHP_UPDATER_PROVIDER',
            'NATIVEPHP_UPDATER_ENABLED',
            'GITHUB_OWNER',
            'GITHUB_REPO',
            'GITHUB_TOKEN',
            'GOOGLE_ANALYTICS_ID',
            'NIGHTWATCH_TOKEN',
            'VITE_APP_SENTRY_VUE_DSN',
            'SENTRY_LARAVEL_DSN',
            'SENTRY_RELEASE',
        ]),
    );
    
    // 👇 ПРОВЕРЯЕМ, ЧТО КЛЮЧ ПОЯВИЛСЯ В .env
    if (file_exists('.env')) {
        $envContent = file_get_contents('.env');
        if (preg_match('/^APP_KEY=(.+)$/m', $envContent, $matches)) {
            echo "✅ APP_KEY successfully set: " . substr(trim($matches[1]), 0, 20) . "...\n";
        } else {
            echo "⚠️ APP_KEY not found in .env file!\n";
        }
    }
}

if (realpath($_SERVER['SCRIPT_FILENAME'] ?? '') === __FILE__) {
    runPrepareReleaseEnvironmentFile();
}