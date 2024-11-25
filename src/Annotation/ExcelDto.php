<?php

declare(strict_types=1);

namespace Firezihai\Excel\Annotation;

/**
 * @Annotation
 * @Target("CLASS")
 */
#[\Attribute(\Attribute::TARGET_CLASS)]
class ExcelDto
{
    /**
     * @param string $type 表格解析类型,title按表头名称，index按表头索引
     * @param int $height 行高
     * @param string $align 文字对齐方式,align:中心居中,left:靠左居中,right:靠右居中
     */
    public function __construct(
        public string $type = 'name',
        public int $height = 45,
        public string $align = 'center'
    ) {}
}
