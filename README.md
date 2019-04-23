# view-office

#### 介绍
使用jacob，poi来实现在线访问docx，doc,xls(支持但显示效果不好),xlsx,ppt，pptx的文件

#### 软件原理
jacob只支持window服务器,使用jacob对word，ppt文件转换为pdf文件来在线查看，使用poi读取数据，再显示再html


#### 安装教程

0. 自行安装office软件
1. 下载该项目
2. 使用idea编译该spring boot项目
3. 执行 java -jar demo.jar

#### 使用说明

1. 调用接口：http://127.0.0.1/perview/office?src=目标文件请求路径

2. word测试请求：http://127.0.0.1/sample/perview/word
![输入图片说明](https://images.gitee.com/uploads/images/2019/0419/173459_cf38794c_1438804.jpeg "SharedScreenshot2.jpg")
3. excel测试请求：http://127.0.0.1/sample/perview/excel
![输入图片说明](https://images.gitee.com/uploads/images/2019/0419/173516_228c806f_1438804.jpeg "SharedScreenshot.jpg")
4. ppt测试请求：http://127.0.0.1/sample/perview/ppt

### 捐赠
![输入图片说明](https://images.gitee.com/uploads/images/2019/0422/085855_688bc575_1438804.jpeg "IMG_0114.JPG")
