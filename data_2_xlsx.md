# 需求文档
## 任务
给你一个 json 文件，帮我写一段 python3 代码，帮我把 json 文件转成 xlsx 表格文件。

## 具体需求
1. 提取json格式中的"data"对象，然后提取"data"对象中的"data_string"元素，data_string 是个 string 类型。
2. data_string 中存在\" 转义符。需要去掉转义符，去掉转义符后，也是个 json 对象，暂且叫它 data_string_json。
3. data_string_json 中取data.series，这是一个数组，数组中每个元素是个json对象，暂且叫它 item_json
4. item_json中有很多字段，每个字段对应表格中的一列。
5. 表格的第一行的每一列，都是字段的英文原名，第二行中，每列要展示对应的中文名称。英文和中文的映射关系在下面。
6. 将表格存储在我指定的文件夹中


# 字段映射关系
timeMinute：时刻
commentCnt：评论数
commentCntRank：互动高峰
commentUcnt：评论人数
consumeUcnt：送礼人数
earnScore：音浪
earnScoreRank：送礼高峰
followUcnt：关注人数
likeCnt：点赞数
likeUcnt：点赞人数
pcuTotal：在线人数
watchUcnt：进入直播间人数
watchUcntRank：在线观众高峰

