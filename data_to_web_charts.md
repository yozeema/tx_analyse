# 需求文档
## 任务
我需要你给我写一个web网页，在网页里，我可以拖入或者上传一个xlsx表格，然后你根据我这个表格，帮我生成echarts趋势图。

## 具体需求
1. 给你的xlsx中，有若干列，其中有一列是timeMinute：时刻，这一列要作为趋势图的横轴。
2. 纵轴展示其他列的内容，要求整个趋势图是折线图形式。
3. 折线图默认展示watchUcnt：进入直播间人数和watchUcntRank：在线观众高峰这两个字段，其他字段默认不展示。
4. 注意，c：主C，music：音乐这两个字段比较特殊，需要在对应的时刻点位上标记出来，拼接在一起展示成“主C：xxx，音乐：xxx”


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
c：主C
music：音乐
memo：备注
