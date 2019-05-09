# ConsoleShellApplication
Window Console Shell Application

### Model
    ppt.exe   Power Point file controller, support demo mode or edit mode  
    wis.exe   Window Input Simulator, support keyboard and mouse, plan to add touch in the future
    wmi.exe   Common Information Model(CIM) management class, management class is WMI class
    
### Support
    (.NET Framework 4)
    All support network(tcp/udp) and process input/output    
    All support process sleep
    ALl support continuous input and mulit command queue
    All support output json and xml format
    
### Example
```
ppt --start       //exit use --stop
ppt --start 2000  //exit use --stop
ppt --run         //exit use --quit
ppt -np -sp 1000 -gp 3 -sp 1000     // ... auto control
ppt -of json -i   //output json format for current ppt file infomation 
```

### 参数
#### ppt(PowerPointShell) 
#### \[ -? | -a | -aw | -np | -pp | -gp | -fp | -lp | -gc | -exp | -exps | -od | -cd | -esv | -qa | -i | --run | --quit | --start | --stop | -of | -sp | -v \]
* -?  显示帮助，与键入 -? -help 是一样的
* -a  激活应用窗体，将活动的演示文档窗体在最前面显示
* -aw(int index)  编辑模式下有效，激活多个编辑文档中的指定一个窗体，参考：-i，示例：-aw 1
* -np 显示文档下一页，示例：-np 或 -of json -np -i
* -pp 显示文档上一页
* -gp(int index)     显示文档指定页面(参考：-i)。
                              示例：-gp 2 或 -of json -gp 2 -i
* -fp 显示文档第一页
* -lp 显示文档最后一页
* -gc(int index)     演示模式下有效，控制动画播放的索引(参考：-i)。
                              示例：-gc 2
*  -exp\[path, width, height\] 将 PowerPoint 文档以图片导出(注意文件完整路径使用'\')。
                           示例：-exp "D:\ttp\" 1920 1080
*  -exps     (index, \[path, width, height\])将 PowerPoint 文档指定的页面以图片导出(注意文件完整路径使用'\')。
                                 示例：-exps 1 "D:\ttp\" 1920 1080
* -od       (path)          打开 PowerPoint 文档，示例：-od "D:\ppt.pptx"
* -cd 关闭活动的 PowerPoint 文档
* -esv 退出活动的 PowerPoint 文档，只在演示模式下有效
* -qa 退出 PowerPoint 应用程序，即关闭所有活动文档
* -i 输出打开的文档/页面信息
* --run 进入持续输入模式
* --quit 退出持续输入模式
* --start\[int port\] 进入网络模式下持续输入模式，端口号空时使用默认端口
* --stop 停止网络模式下持续输入模式
* -of(enum)          输出格式：0:default默认格式，1:json格式，2:xml格式。示例：-of 1 或 -of json
* -sp(int ms)        将当前线程挂起指定的时间(ms)
* -v 控制台程序版本信息

#### wis(WindowsInput) 
#### \[ -? | -kd | -ku | -kp | -ks | -te | -kvs | -mld | -mlu | -mlc | -mldc | -mrd | -mru | -mrc | -mrdc | -mmb | -mmt | -hs | -vs | -mmtvd | -gmp | --run | --quit | --start | --stop | -of | -sp | -v \]
* -? 显示帮助，与键入 -? -help 是一样的
* -kd(enum)          模拟输入键盘按下键，参数为十进制键值或是键的枚举字符串，参考：-kvs
* -ku(enum)          模拟输入键盘释放键
* -kp(enum)          模拟输入键盘按压并释放键
* -ks(keys)          模拟输入组合键。
                              例如：-ks CONTROL+VK_C 或 -ks CONTROL|LSHIFT+VK_A
* -te(str)           模拟输入文本字符
* -kvs 键盘键值参考
* -mld 模拟输入鼠标左键按下
* -mlu 模拟输入鼠标左键弹起
* -mlc 模拟输入鼠标左键点击
* -mldc 模拟输入鼠标左键双击
* -mrd 模拟输入鼠标右键按下
* -mru 模拟输入鼠标右键弹起
* -mrc 模拟输入鼠标右键点击
* -mrdc 模拟输入鼠标右键双击
* -mmb(point)         模拟输入鼠标坐标相对偏移。
                              示例：-mmb 100,100
* -mmt(point)         模拟输入鼠标坐标移动到绝对位置。
                              示例：-mmt 100,100
* -hs(int)           模拟输入水平滚动
* -vs(int)           模拟输入垂直滚动
* -mmtvd(point)         模拟输入移动鼠标位置到虚拟桌面
* -gmp 获取鼠标相对屏幕的坐标
* --run 进入持续输入模式
* --quit 退出持续输入模式
* --start[int port] 进入网络模式下持续输入模式，端口号空时使用默认端口
* --stop 停止网络模式下持续输入模式
* -of(enum)          输出格式：0:default默认格式，1:json格式，2:xml格式。示例：-of 1 或 -of json
* -sp(int ms)        将当前线程挂起指定的时间(ms)
* -v 控制台程序版本信息

#### wmi(WMIShell) 
#### \[ -? | -ex | -li | -mos | -mosp | -mosi | --run | --quit | --start | --stop | -of | -sp | -v \]
* -? 显示帮助，与键入 -? -help 是一样的
* -ex 列举内置默认的信息查询语句
* -li 列举 Win32 可查询的信息管理对象(表)
* -mos(str)           Management Object Searcher 管理信息的指定查询；参数：SQL 语句，参考表：-li -ex
* -mosp(enum)          Management Object Searcher 管理信息的字段属性查询；参数参考：-li，使用索引或名称
* -mosi(int)           Management Object Searcher 管理信息的指定查询；参数：-ex 的索引，引用内置示例查询
* --run 进入持续输入模式
* --quit 退出持续输入模式
* --start[int port] 进入网络模式下持续输入模式，端口号空时使用默认端口
* --stop 停止网络模式下持续输入模式
* -of(enum)          输出格式：0:default默认格式，1:json格式，2:xml格式。示例：-of 1 或 -of json
* -sp(int ms)        将当前线程挂起指定的时间(ms)
* -v 控制台程序版本信息

### Declare
    There may be some bugs ←←@_@
