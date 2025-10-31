<?php

declare(strict_types=1);

namespace Firezihai\Excel\Annotation;

/**
 * @Annotation
 * @Target("PROPERTY")
 */
#[\Attribute(\Attribute::TARGET_PROPERTY)]
class ExcelHeader
{
    /**
     * @param string $name 表头名称
     * @param int $index 表头索引
     * @param int $width 列宽
     * @param string $source 数据来源映射,导出时使用，例customer_name,实际对应的是customer.name
     * @param string $align 文字对齐方式,align:中心居中,left:靠左居中,right:靠右居中
     * @param string $color 表头文字颜色
     * @param bool $formatter 是否开启格式化,当值为true时，须要定义以formatter开头的自定义格式化方法，例如,status字段开启格式化
     *                        必须定义formatterStatus方法，对status字段格式化
     */
    public function __construct(
        public string $name,
        public ?int $index = null,
        public ?int $width = 20,
        public ?string $source = null,
        public ?string $align = 'center',
        public ?string $color = null,
        public mixed $formatter = null
    ) {}
}
