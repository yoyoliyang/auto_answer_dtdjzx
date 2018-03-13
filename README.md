# auto_answer_dtdjzx
灯塔在线自动答题脚本（PC,android）

流程：

下载randomList100遍，然后利用re来确定标题、答案，并生成一个标题：答案的题库字典dic{}。

读取user.xlsx中的用户和密码，自动填写到登录界面。

正式答题的时候，抓取getGameSubject，利用re获取20个标题（subject），然后在字典题库中找到对应答案（如果在字典中找不到，那么会提示，之后手动加入到字典文件中dic.txt）。

答每道题的时候，利用opencv的模板匹配，找到abcd四个选项的具体坐标，进行点击。

更多的东西看代码吧，本人刚开始接触python，代码肯定是不够精简。
