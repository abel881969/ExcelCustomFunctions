# ExcelCustomFunctions

[English](./README.md)

基于微软高级公式环境自定义 Excel 函数

---

## 准备工作

函数基于 LAMBDA 函数，请保持你的 Excel 处于最新版本，以便正常使用 LAMBDA 函数。具体来说，Excel 应不低于以下版本：

- Windows: 16.0.14729.20260
- Mac: 16.56 (Build 21121100)
- iOS: 2.56 (Build 21120700)
- Android: 16.0.14729.20176

安装新的名为"Advanced formula environment, a Microsoft Garage project"的加载项，可以很容易地导入/导出和编辑自定义LAMBDA函数。

访问以下链接可以获得更多信息：

- [Announcing LAMBDAs to Production and Advanced Formula Environment, A Microsoft Garage Project
](https://techcommunity.microsoft.com/t5/excel-blog/announcing-lambdas-to-production-and-advanced-formula/ba-p/3073293)
- [Big News！这个函数正式发布了，远远超出期待-LAMBDA](https://mp.weixin.qq.com/s/j3C7MQodnt9KbROWos4fkw)

## 使用函数

只需要简单地导入文本文件，就可以使用自定义函数。

在项目中选择需要的函数列表，在 Advanced Formula Environment 中导入即可，可以通过下载文本文件或直接复制链接导入。

- 统计相关函数：./statistics
- 金融相关函数：./finance
- 财务相关函数：./accounting
- 日期相关函数：./datetime

## 共享函数，助力项目

如果你愿意共享你的函数或者创意，请通过 PR 的方式提供给我们，如果你有定制函数的需求，也可以通过 issue 提出，我们会考虑提供支持。

### 函数提交规范

- 自定义函数应包含注释，明确函数名称、函数简介、参数及参数定义、示例和函数本体
- 通用函数名称宜采用英文大写字母，与 Excel 原生函数风格保持一致。特定功能的函数宜采用驼峰命名，简洁、准确的描述函数功能
- 参数命名宜采用英文小写字母，多个单词以下划线分割，与 Excel 原生函数参数风格保持一致。参数应注明参数含义和对类型的特殊要求
- 案例应简洁、清晰、全面覆盖函数的功能
- 函数本体宜简洁、准确，避免无意义代码

提交示例:

```
/*

函数名称: IFBLANK

函数简介: 判断值是否为空，如果为空返回特定值

参数:

value: 需要判断是否为空的值

value_if_blank: 如果判断为真的返回值

示例:

=IFBLANK(,"blankVal")

*/

IFBLANK =LAMBDA(value, value_if_blank, IF(ISBLANK(value),value_if_blank,value));
```

### 创意提交规范

待完善

### 需求提交规范

待完善