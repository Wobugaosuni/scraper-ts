# 爬取 Indiegogo 数据

1. 用户自定义需要爬取的网页地址列表 `urlList`
2. 由于网页部分dom是通过Js加载，使用puppeteer模拟浏览器打开，并获取到网页内容
3. 遍历 `urlList`，获取表格需要的内容
4. 生成表格
5. 复制表格内容到在线excel文档