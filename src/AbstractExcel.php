<?php

declare(strict_types=1);

namespace Firezihai\Excel;

use Firezihai\Excel\Annotation\ExcelDto;
use Firezihai\Excel\Annotation\ExcelHeader;

abstract class AbstractExcel
{
    protected $dto;

    protected $dtoAnnotion = ExcelDto::class;

    protected $dtoPropertyAnnotion = ExcelHeader::class;

    /**
     * 解析注解配置.
     * @return array
     */
    public function parseAnnotation(string $class)
    {
        $reflection = new \ReflectionClass($class);
        $result = $this->getAttributes($reflection);
        $annotations = [];
        foreach ($result as $ref) {
            $annotations['_c'][get_class($ref)] = $ref;
        }
        $props = $reflection->getProperties(\ReflectionProperty::IS_PUBLIC | \ReflectionProperty::IS_PROTECTED);
        foreach ($props as $prop) {
            $attrs = $this->getAttributes($prop);
            foreach ($attrs as $atrr) {
                $annotations['_p'][$prop->name][get_class($atrr)] = $atrr;
            }
        }
        return $annotations;
    }

    /**
     * 获取类属生
     * @return null[]
     * @throws \InvalidArgumentException
     */
    public function getAttributes(\Reflector $reflection)
    {
        $result = [];
        $attributes = $reflection->getAttributes();
        foreach ($attributes as $attribute) {
            if (! class_exists($attribute->getName())) {
                $className = $methodName = $propertyName = '';
                if ($reflection instanceof \ReflectionClass) {
                    $className = $reflection->getName();
                } elseif ($reflection instanceof \ReflectionMethod) {
                    $className = $reflection->getDeclaringClass()->getName();
                    $methodName = $reflection->getName();
                } elseif ($reflection instanceof \ReflectionProperty) {
                    $className = $reflection->getDeclaringClass()->getName();
                    $propertyName = $reflection->getName();
                }
                $message = sprintf(
                    "No attribute class found for '%s' in %s",
                    $attribute->getName(),
                    $className
                );
                if ($methodName) {
                    $message .= sprintf('->%s() method', $methodName);
                }
                if ($propertyName) {
                    $message .= sprintf('::$%s property', $propertyName);
                }
                throw new \InvalidArgumentException($message);
            }
            $result[] = $attribute->newInstance();
        }
        return $result;
    }

    /**
     * 表头排序.
     */
    public function sortHeader(array $excelHeader)
    {
        $newHeader = [];
        foreach ($excelHeader as $value) {
            if ($value['index']) {
                // 对应列已设置表头，把表头插到原表头的前面
                if (isset($newHeader[$value['index']])) {
                    $newHeader = $this->arrayPrepend($newHeader, $value['index'], $value);
                } else {
                    $newHeader[$value['index']] = $value;
                }
            } else {
                $newHeader[] = $value;
            }
        }
        // 按索引排序
        ksort($newHeader);
        return $newHeader;
    }

    /**
     * 将元素添加到索引前面.
     */
    public function arrayPrepend(array $array, string $value, ?string $key = null)
    {
        if (is_null($key)) {
            array_unshift($array, $value);
        } else {
            $array = [$key => $value] + $array;
        }
    }

    /**
     * 获取dto类表注解.
     * @throws \InvalidArgumentException
     */
    protected function parseDtoAnnotation(string $dto): array
    {
        $annotationMate = $this->parseAnnotation($dto);

        if (empty($annotationMate) || ! isset($annotationMate['_c'])) {
            throw new \InvalidArgumentException('dto annotation info is empty');
        }

        $type = $annotationMate['_c'][$this->dtoAnnotion]->type ?? 'name';
        $result = [];
        foreach ($annotationMate['_c'][$this->dtoAnnotion] as $name => $mate) {
            $result[$name] = $mate;
        }
        $header = [];
        foreach ($annotationMate['_p'] as $name => $mate) {
            $item['field'] = $name;
            foreach ($mate[$this->dtoPropertyAnnotion] as $name => $mate) {
                $item[$name] = $mate;
            }

            $header[] = $item;
        }
        $result['header'] = $header;
        return $result;
    }

    /**
     * 获取 excel 列索引.
     */
    protected function getColumnIndex(int $columnIndex = 0): string
    {
        if ($columnIndex < 26) {
            return chr(65 + $columnIndex);
        }
        if ($columnIndex < 702) {
            return chr(64 + intval($columnIndex / 26)) . chr(65 + $columnIndex % 26);
        }
        return chr(64 + intval(($columnIndex - 26) / 676)) . chr(65 + intval((($columnIndex - 26) % 676) / 26)) . chr(65 + $columnIndex % 26);
    }

    /**
     * @param array $item 单行数据
     * @param array $header 字段导出表头配置
     */
    protected function getFieldValue(array $item, array $header)
    {
        $formaterMethod = $header['formatter'] ? $this->dto . '::formatter' . ucfirst($header['field']) : '';

        $value = isset($item[$header['field']]) ? $item[$header['field']] : '';

        if ($header['source'] && ! str_contains($header['source'], '.')) {
            $value = isset($item[$header['source']]) ? $item[$header['source']] : '';
        }

        if ($header['source'] && str_contains($header['source'], '.')) {
            $arrKeys = explode('.', $header['source']);
            $value = $item;
            foreach ($arrKeys as $k) {
                if (isset($value[$k])) {
                    $value = $value[$k];
                } else {
                    $value = '';
                    break;
                }
            }
        }

        return $formaterMethod ? $formaterMethod($value) : $value;
    }

    protected function getTmpDir(): string
    {
        $tmp = ini_get('upload_tmp_dir');

        if ($tmp !== false && file_exists($tmp)) {
            return realpath($tmp) . '/';
        }

        return realpath(sys_get_temp_dir()) . '/';
    }
}
