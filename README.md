# Excel

## 功能

- 用于解析表格内容
- 支持phpoffice、xlswriter


## 安装

```shell
composer require firezihai/excel -vvv

```

## 使用

### 定义DTO

```php
//index按表头索引解析表格，name：按表头名称解析表格，默认name
#[ExcelDto(type:"index")]
class UserDto implements ExcelDtoInterface
{
    // 导出表格无index时，属性在类中的顺序，即为表头的顺序
    #[ExcelHeader(name:"用户名",index:0)]
    public string $username;
    
    #[ExcelHeader(name:"昵称",index:1)]
    public string $nickname;
    
    #[ExcelHeader(name:"出生日期",index:3)]
    public string $birthday;
    
    #[ExcelHeader(name:"性别",index:2)]
    public string $gender;
}

```
### 解析表格

```php
$excelFactory = new ExcelFactory();
$excel = $excelFactory->get('phpOffice');
$data =  $excel->parse('./test.xlsx', UserDto::class);

```
> 注意 `xlswriter` 解析表格时，需安装 [xlswriter](https://github.com/viest/php-ext-xlswriter) 扩展

