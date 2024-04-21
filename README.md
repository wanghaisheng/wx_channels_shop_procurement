# 视频号订单补货统计工具
看着一起干活的小妹姐补货统计太花时间了，于是花了一点点时间写出了这个程序。
这是一个用于统计视频号订单的Go语言程序。它可以读取视频号订单导出的Excel文件,把订单中相同的商品SKU统计分组并生成统计结果。
## 运行示例
1.拖动导出的订单excel到exe文件上就能得到excel同名文件的txt结果文件。
2.exe文件请到releases下载default.zip。
3.可能会被杀毒杀掉。- - 把exe文件添加到排除区。

<img width="1189" alt="image" src="https://github.com/ruan11223344/wx_channels_shop_procurement/assets/5679023/37a66cfa-5cea-4be8-8b86-8e3ab5ad2663">

<img width="1186" alt="image" src="https://github.com/ruan11223344/wx_channels_shop_procurement/assets/5679023/a7652ce8-7f4e-46bc-a28d-4c49f32f43f7">
## 特点


- 自动统计商品数量:程序会根据商品名称和规格自动统计每个商品的数量。
- 处理备注订单:对于包含买家备注或商家备注的订单,程序会单独列出,不计入统计结果。
- 生成txt文件:程序会在可执行文件所在目录下生成一个与Excel文件同名的txt文件,包含统计结果和备注订单信息。
- 方便使用:只需将视频号小店订单导出的Excel文件拖放到程序上,即可自动生成统计结果。

## 使用方法

1. 将视频号小店订单导出为Excel文件。- 订单/配送 -> 订单管理 -> 筛选订单后 -> 导出订单  
2. 将Excel文件拖放到本程序的exe文件上即可。
3. 程序会自动读取Excel文件,生成统计结果,并在程序所在目录下生成一个与Excel文件同名的txt文件。
4. 打开生成的txt文件,查看统计结果和备注订单信息。

## 统计结果格式

统计结果包括两部分:商品统计(不包含买家备注跟卖家备注的订单)和备注订单。

### 商品统计

每个商品的统计信息包括:

- 商品名称
- 商品价格
- 每个SKU的数量

示例:

```
1.商品A 商品价格:100

SKU:红色 数量:10

SKU:蓝色 数量:5
```

### 备注订单

备注订单列出了包含买家备注或商家备注的订单,不计入商品统计。每个备注订单包括:

- 商品名称
- 原始SKU
- 单价
- 买家备注(如果有)
- 商家备注(如果有)
- 收件人地址
- 省、市、区
- 收件人手机号

示例:

```
额外的备注订单(未统计在上方):

1.商品B 原始SKU:L 单价:200 买家备注:请尽快发货 收件人地址:XXX 省:XX 市:XX 区:XX 收件人手机:1234567890
```

## 注意事项

- 程序依赖于Excel文件的格式,请确保导出的Excel文件格式正确。
- 程序会覆盖已有的同名txt文件,请注意备份。

## 依赖库

- [github.com/xuri/excelize/v2](https://github.com/xuri/excelize)

请确保在运行程序前安装了以上依赖库。

## 许可证

本项目采用 [MIT 许可证](LICENSE)。