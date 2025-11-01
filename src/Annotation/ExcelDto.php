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
     * @param string $type 表格解析类型,name按表头名称，index按表头索引
     * @param int $height 行高
     * @param string $align 文字对齐方式,align:中心居中,left:靠左居中,right:靠右居中
     * @param bool $checkHeader 检查表头的名称是否和注解里配置的表头名称一样,只在type = name下有效
     */
    public function __construct(
        public string $type = 'name',
        public int $height = 45,
        public string $align = 'center',
        public bool $checkHeader = false
    ) {}
}
