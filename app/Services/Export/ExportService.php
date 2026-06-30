<?php

declare(strict_types=1);

namespace App\Services\Export;

use App\Enums\ExportColumnEnum;
use App\Enums\TimestampTypeEnum;
use App\Models\Project;
use App\Models\Timestamp;
use App\Services\HolidayService;
use App\Services\LocaleService;
use App\Services\TimestampService;
use App\Settings\ExportSettings;
use Carbon\Carbon;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\Date;
use Illuminate\Support\Str;
use OpenSpout\Common\Entity\Style\Style;
use Spatie\LaravelPdf\Facades\Pdf;
use Spatie\SimpleExcel\SimpleExcelWriter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;

class ExportService
{
    private readonly ?Carbon $startDate;

    private readonly ?Carbon $endDate;

    public function __construct(private readonly array $timestampTypes, ?string $startDate, ?string $endDate, private readonly array $projectIds = [])
    {
        new LocaleService;
        $this->startDate = $startDate ? Date::parse($startDate)->startOfDay() : null;
        $this->endDate = $endDate ? Date::parse($endDate)->endOfDay() : null;
    }

    public function generateFileName(string $extension): string
    {
        $exportFileName = 'TimeScribe-Export';
        if ($this->startDate && $this->endDate) {
            $exportFileName .= ' — '.$this->startDate->format('Y-m-d').' - '.$this->endDate->format('Y-m-d');
        }
        if ($this->projectIds) {
            $projectNames = Project::withTrashed()->whereIn('id', $this->projectIds)->get('name')->map(fn ($projectName) => Str::slug($projectName['name']))->join(' # ');
            $exportFileName .= ' — '.$projectNames;
        }

        return $exportFileName.'.'.$extension;
    }

    private function getExportData(): Collection
    {
        $timestamps = Timestamp::query()->with(['project']);
        $timestamps->whereIn('type', $this->timestampTypes);
        if ($this->startDate) {
            $timestamps->where('started_at', '>=', $this->startDate);
        }
        if ($this->endDate) {
            $timestamps->where('ended_at', '<=', $this->endDate);
        }
        if ($this->projectIds) {
            $timestamps->whereIn('project_id', $this->projectIds);
        }

        $timestamps = $timestamps->latest('started_at')->get();
        $additionalTimestamps = $this->getSickLeaveHolidayData($timestamps->last()?->started_at);

        return collect([...$timestamps, ...$additionalTimestamps])->sortByDesc('started_at')->values();
    }

    private function getSickLeaveHolidayData(?Carbon $firstDate = null): Collection
    {
        $dateList = collect();
        $startDate = $this->startDate ?? $firstDate;
        $endDate = $this->endDate ?? now()->endOfDay();

        $absences = TimestampService::getAbsence($startDate, $endDate);
        if (in_array(TimestampTypeEnum::SICK->value, $this->timestampTypes)) {
            $dateList = $dateList->merge($absences->where('type', 'sick')->map(fn ($item): array => [
                'date' => $item->date,
                'type' => 'sick',
                'duration' => 0,
            ]));
        }

        if (in_array(TimestampTypeEnum::HOLIDAY->value, $this->timestampTypes)) {
            $yearRange = range($startDate->year, $endDate->year);
            $holidays = HolidayService::getHoliday($yearRange)->filter(fn ($date): bool => $date <= $endDate && $date >= $startDate);

            $dateList = $dateList->merge($holidays->map(fn ($item): array => [
                'date' => $item,
                'type' => 'holiday',
                'duration' => 0,
            ]));
        }

        if (in_array(TimestampTypeEnum::VACATION->value, $this->timestampTypes)) {
            $dateList = $dateList->merge($absences->where('type', 'vacation')->map(fn ($item): array => [
                'date' => $item->date,
                'type' => 'vacation',
                'duration' => 0,
            ]));
        }

        return $dateList->unique('date')->map(function (array $item) {
            $timestampStartedAt = $item['date'];
            $timestampEndedAt = $item['date']->clone()->addSeconds(TimestampService::getPlan($item['date']) * 3600);

            return Timestamp::make([
                'started_at' => $timestampStartedAt,
                'created_at' => $timestampStartedAt,
                'ended_at' => $timestampEndedAt,
                'updated_at' => $timestampEndedAt,
                'last_ping_at' => $timestampEndedAt,
                'type' => $item['type'],
            ]);
        })->sortBy('started_at')->values();
    }

    public function exportAsCsv(string $filePath): void
    {
        $file = fopen($filePath, 'w');
        fputcsv($file, $this->headerArray(), escape: '\\');

        foreach ($this->getExportData() as $timestamp) {
            fputcsv($file, $this->timestampToRowArray($timestamp), escape: '\\');
        }

        fclose($file);
    }

    public function exportAsExcel(string $filePath): void
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        // Устанавливаем шрифт Times New Roman 12pt для всего документа
        $sheet->getParent()->getDefaultStyle()->getFont()->setName('Times New Roman')->setSize(12);
        
        // Устанавливаем ширину столбцов в пикселях (1px = 0.75 символа в PhpSpreadsheet)
        $sheet->getColumnDimension('A')->setWidth(62 * 0.75);
        $sheet->getColumnDimension('B')->setWidth(417 * 0.75);
        $sheet->getColumnDimension('C')->setWidth(417 * 0.75);
        $sheet->getColumnDimension('D')->setWidth(417 * 0.75);
        $sheet->getColumnDimension('E')->setWidth(109 * 0.75);
        $sheet->getColumnDimension('F')->setWidth(109 * 0.75);
        $sheet->getColumnDimension('G')->setWidth(109 * 0.75);
        
        $rowIndex = 1;
        
        // Строка 1: Заголовок "Программист"
        $sheet->mergeCells('A' . $rowIndex . ':G' . $rowIndex);
        $sheet->setCellValue('A' . $rowIndex, 'Программист');
        $sheet->getStyle('A' . $rowIndex)->applyFromArray([
            'font' => [
                'bold' => true,
                'name' => 'Times New Roman',
                'size' => 12,
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'B7DEE8'],
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
        ]);
        $rowIndex++;
        
        // Получаем данные
        $exportData = $this->getExportData();
        
        // Группируем записи по проектам
        $groupedData = [];
        foreach ($exportData as $timestamp) {
            $projectName = $timestamp->project?->name ?? 'Без проекта';
            if (!isset($groupedData[$projectName])) {
                $groupedData[$projectName] = [];
            }
            $groupedData[$projectName][] = $timestamp;
        }
        
        // Проходим по каждой группе проектов
        foreach ($groupedData as $projectName => $timestamps) {
            // Строка с названием проекта
            $sheet->mergeCells('A' . $rowIndex . ':G' . $rowIndex);
            $sheet->setCellValue('A' . $rowIndex, $projectName);
            $sheet->getStyle('A' . $rowIndex)->applyFromArray([
                'font' => [
                    'bold' => true,
                    'name' => 'Times New Roman',
                    'size' => 12,
                ],
                'fill' => [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'startColor' => ['rgb' => 'D9D9D9'],
                ],
                'alignment' => [
                    'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                    'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                ],
            ]);
            $rowIndex++;
            
            // Записи отработанных часов
            $counter = 1;
            foreach ($timestamps as $timestamp) {
                // Столбец A: Нумерация
                $sheet->setCellValue('A' . $rowIndex, $counter);
                
                // Столбец B: Дата в формате "24 июня 2026 года"
                $date = $timestamp->started_at;
                $months = [
                    1 => 'января', 2 => 'февраля', 3 => 'марта',
                    4 => 'апреля', 5 => 'мая', 6 => 'июня',
                    7 => 'июля', 8 => 'августа', 9 => 'сентября',
                    10 => 'октября', 11 => 'ноября', 12 => 'декабря'
                ];
                $formattedDate = $date->day . ' ' . $months[$date->month] . ' ' . $date->year . ' года';
                $sheet->setCellValue('B' . $rowIndex, $formattedDate);
                
                // Столбец C: Длительность в формате "H ч. MM мин."
                $duration = $timestamp->duration;
                $hours = floor($duration / 3600);
                $minutes = floor(($duration % 3600) / 60);
                $formattedDuration = $hours . ' ч. ' . $minutes . ' мин.';
                $sheet->setCellValue('C' . $rowIndex, $formattedDuration);
                
                // Столбец D: "Человеческие и временные"
                $sheet->setCellValue('D' . $rowIndex, 'Человеческие и временные');
                
                // Столбец E: пусто
                $sheet->setCellValue('E' . $rowIndex, '');
                
                // Столбец F: "отработанное время" с выравниванием по центру
                $sheet->setCellValue('F' . $rowIndex, 'отработанное время');
                $sheet->getStyle('F' . $rowIndex)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                
                // Столбец G: пусто
                $sheet->setCellValue('G' . $rowIndex, '');
                
                $counter++;
                $rowIndex++;
            }
        }
        
        // Сохраняем файл
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($filePath);
    }

    public function exportAsPdf(string $filePath): void
    {
        $exportData = $this->getExportData();
        $workTime = $exportData->where('type', 'work')->sum('duration');
        $breakTime = $exportData->where('type', 'break')->sum('duration');
        $sickTime = $exportData->where('type', 'sick')->sum('duration');
        $vacationTime = $exportData->where('type', 'vacation')->sum('duration');
        $holidayTime = $exportData->where('type', 'holiday')->sum('duration');

        $totalHours = floor($workTime / 3600);
        $totalMinutes = floor(($workTime % 3600) / 60);
        $totalFormatted = sprintf('%d:%02d', $totalHours, $totalMinutes);

        $exportSettings = resolve(ExportSettings::class);

        $projects = [];
        if ($this->projectIds) {
            $projects = Project::withTrashed()->whereIn('id', $this->projectIds)->get(['name', 'color']);
        }

        Pdf::view('pdf.export', [
            'timestamps' => $exportData->map(fn (Timestamp $timestamp): array => $this->timestampToRowArray($timestamp, true))->all(),
            'columns' => $this->headerArray(),
            'startDate' => $this->startDate?->isoFormat('L'),
            'endDate' => $this->endDate?->isoFormat('L'),
            'projects' => $projects,
            'totalHours' => $totalFormatted,
            'breakTime' => $breakTime,
            'workTime' => $workTime,
            'sickTime' => $sickTime,
            'vacationTime' => $vacationTime,
            'holidayTime' => $holidayTime,
        ])
            ->format($exportSettings->pdf_paper_size)
            ->orientation($exportSettings->pdf_orientation)
            ->save($filePath);
    }

    private function headerArray(): array
    {
        return collect(ExportColumnEnum::toResource())->filter(fn ($column) => $column['is_visible'])->map(fn ($column) => $column['label'])->toArray();
    }

    private function timestampToRowArray(Timestamp $timestamp, bool $isoFormat = false): array
    {
        $all = [
            'type' => $timestamp['type']->value,
            'description' => $timestamp['description'] ?? '',
            'metadata' => $timestamp['project']?->metadata ?? '',
            'project' => $timestamp['project'] ? implode(' ', [$timestamp['project']->icon, $timestamp['project']->name]) : '',
            'import_source' => $timestamp['source'] ?? '',
            'duration' => $timestamp['ended_at'] ? gmdate('H:i:s', (int) $timestamp['started_at']->diffInSeconds($timestamp['ended_at'])) : '',
            'hourly_rate' => $timestamp['project']?->hourly_rate ? number_format($timestamp['project']->hourly_rate, 2) : '',
            'billable_amount' => $timestamp['duration'] && $timestamp['project']?->hourly_rate ? number_format($timestamp['duration'] / 60 * $timestamp['project']?->hourly_rate / 60, 2) : '',
            'currency' => $timestamp['project']?->hourly_rate ? $timestamp['project']?->currency ?? '' : '',
            'paid' => $timestamp['paid'] ? 'Yes' : '',
        ];

        if ($isoFormat) {
            $all['start_date'] = $timestamp['started_at']->isoFormat('L');
            $all['start_time'] = $timestamp['started_at']->isoFormat('LTS');
            $all['end_date'] = $timestamp['ended_at'] ? $timestamp['ended_at']->isoFormat('L') : '';
            $all['end_time'] = $timestamp['ended_at'] ? $timestamp['ended_at']->isoFormat('LTS') : '';
        } else {
            $all['start_date'] = $timestamp['started_at']->format('d/m/Y');
            $all['start_time'] = $timestamp['started_at']->format('H:i:s');
            $all['end_date'] = $timestamp['ended_at'] ? $timestamp['ended_at']->format('d/m/Y') : '';
            $all['end_time'] = $timestamp['ended_at'] ? $timestamp['ended_at']->format('H:i:s') : '';
        }

        return collect($this->headerArray())->mapWithKeys(fn ($value, $key): array => [$key => $all[$key] ?? ''])->toArray();
    }
}
