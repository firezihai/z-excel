<?php

declare(strict_types=1);

namespace Firezihai\Excel\Driver;

use Firezihai\Excel\AbstractExcel;
use Firezihai\Excel\ExcelInterface;
use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

class XlsWriter extends AbstractExcel implements ExcelInterface
{
    private $aligns = [
        'left' => [Format::FORMAT_ALIGN_LEFT, Format::FORMAT_ALIGN_VERTICAL_CENTER],
        'center' => [Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER],
        'right' => [Format::FORMAT_ALIGN_RIGHT, Format::FORMAT_ALIGN_VERTICAL_CENTER],
    ];

    public function getExportHeader(string $dto)
    {
        $annotationMeta = $this->parseDtoAnnotation($dto);
        $excelHeader = $this->sortHeader($annotationMeta['header']);
        return $annotationMeta['header'];
    }

    public function parse(string $filename, string $dto)
    {
        $config = ['path' => dirname($filename)];
        $excel = new Excel($config);
        $excel->openFile(basename($filename))->openSheet(null, Excel::SKIP_EMPTY_ROW);
        $annotationMate = $this->parseDtoAnnotation($dto);
        $header = $annotationMate['header'];
        $fieldMap = [];
        $type = $annotationMate['type'] ?? 'name';
        foreach ($header as $value) {
            $fieldMap[$value[$type]] = $value['field'];
        }
        // $header = array_flip($header);
        $fieldCell = [];
        $i = 0;
        $data = [];
        $types = [];
        $excelHeader = [];
        while (($row = $excel->nextRow($types)) !== null) {
            if ($i == 0) {
                foreach ($row as $key => $value) {
                    $fieldKey = $annotationMate['type'] == 'index' ? $key : $value;
                    // 空跳过
                    if (empty($value) || ! isset($fieldMap[$fieldKey])) {
                        continue;
                    }
                    $fieldCell[] = [
                        'index' => $key,
                        'field' => $fieldMap[$fieldKey],
                    ];
                    //  时间格式有点特殊，是一个数值，需要将时间列设置为时间格式
                    if (strpos($value, '时间') !== false || strpos($value, '日期') !== false) {
                        $types[$key] = Excel::TYPE_TIMESTAMP;
                    }
                    $excelHeader[] = $value;
                }
                // 只在按表头名称获取数据时可以检查表头
                if ($type === 'name') {
                    file_put_contents('zhai.txt', join(',', $excelHeader));
                    $this->checkHeader($annotationMate['checkHeader'], $excelHeader, $header);
                }
            } else {
                $temp = [];
                foreach ($fieldCell as $cell) {
                    $value = $row[$cell['index']];

                    // 时间列的时间戳转成日期格式
                    if (isset($types[$cell['index']]) && $types[$cell['index']] == Excel::TYPE_TIMESTAMP) {
                        // 表格中的时间格式为2023/03/01 或者2023-01-0
                        if ($value && is_numeric($value)) {
                            $temp[$cell['field']] = date('Y-m-d H:i:s', $value);
                        // 表格中的时间格式为2023年10月1日 保留原格式
                        } elseif ($value && is_string($value)) {
                            $temp[$cell['field']] = $value;
                        } else {
                            $temp[$cell['field']] = '';
                        }
                    } else {
                        $temp[$cell['field']] = $value;
                    }
                }
                if (! empty($temp)) {
                    $data[] = $temp;
                }
            }

            ++$i;
        }
        return $data;
    }

    public function export(string $dto, array $data, array $exportHeader = [])
    {
        $this->dto = $dto;
        $annotationMeta = $this->parseDtoAnnotation($dto);

        $excelHeader = $this->sortHeader($annotationMeta['header']);
        if (! empty($exportHeader)) {
            $excelHeader = $this->filterHeader($excelHeader, $exportHeader);
        }

        $excel = new Excel(['path' => $this->getTmpDir() . '/']);
        $tempFileName = time() . '.xlsx';
        $fileObject = $excel->fileName($tempFileName);

        // 全局默认样式
        $format = new Format($excel->getHandle());
        if (! empty($annotationMeta['align'])) {
            $align = $this->aligns[$annotationMeta['align']];
            // 不使用...语法
            $format->align($align[0], $align[1]);
        }
        if (! empty($annotationMeta['height'])) {
            $excel->setRow('A1:' . $this->getColumnIndex(count($excelHeader)) . (count($data) + 1), $annotationMeta['height']);
        }
        // header方法调用必须在设置样式之后
        $excel->defaultFormat($format->toResource());

        $this->fillHeader($excel, $annotationMeta, $excelHeader);

        try {
            $row = 1;
            foreach ($data as $item) {
                $column = 0;
                foreach ($excelHeader as $header) {
                    //   $coordinate = $this->getColumnIndex($headerIndex) . $row;
                    $excel->insertText($row, $column, $this->getFieldValue($item, $header));
                    ++$column;
                }
                ++$row;
            }
        } catch (\RuntimeException $e) {
        }

        return $fileObject->output();
    }

    /**
     * 填充表头.
     */
    public function fillHeader(Excel $excel, array $annotationMeta, array $excelHeader)
    {
        $headerIndex = 0;
        foreach ($excelHeader as $item) {
            $headerColumn = $this->getColumnIndex($headerIndex);
            $columnFormat = new Format($excel->getHandle());
            $width = ! empty($item['width']) ? $item['width'] : 40;

            if (! empty($item['color'])) {
                $columnFormat->fontColor(intval(str_replace('#', '0x', $item['color'])));
            }
            if (! empty($item['align'])) {
                $align = $this->aligns[$item['align']];
                // 不使用...语法
                $columnFormat->align($align[0], $align[1]);
            }

            $excel->setColumn($headerColumn . ':' . $headerColumn, $width, $columnFormat->toResource());

            ++$headerIndex;
        }

        $columnName = array_column($excelHeader, 'name');

        $excel->header($columnName);
    }
}
