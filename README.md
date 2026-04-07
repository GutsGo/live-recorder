![video_spider](https://socialify.git.ci/GutsGo/live-recorder/image?font=Inter&forks=1&language=1&owner=1&pattern=Circuit%20Board&stargazers=1&theme=Light)

## 💡简介
[![Python Version](https://img.shields.io/badge/python-3.11.6-blue.svg)](https://www.python.org/downloads/release/python-3116/)
[![Supported Platforms](https://img.shields.io/badge/platforms-Windows%20%7C%20Linux-blue.svg)](https://github.com/GutsGo/live-recorder)
[![Docker Pulls](https://img.shields.io/docker/pulls/imagindragons/live-recorder?label=Docker%20Pulls&color=blue&logo=docker)](https://hub.docker.com/r/imagindragons/live-recorder/tags)
![GitHub issues](https://img.shields.io/github/issues/GutsGo/live-recorder.svg)
[![Latest Release](https://img.shields.io/github/v/release/GutsGo/live-recorder)](https://github.com/GutsGo/live-recorder/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/GutsGo/live-recorder/total)](https://github.com/GutsGo/live-recorder/releases/latest)

一款**简易**的可循环值守的直播录制工具，基于FFmpeg实现多平台直播源录制，支持自定义配置录制以及直播状态推送。

## 😺已支持平台

- [x] 抖音
- [x] TikTok
- [x] 快手
- [x] 虎牙
- [x] 斗鱼
- [x] YY
- [x] B站
- [x] 小红书
- [x] bigo 
- [x] blued
- [x] SOOP(原AfreecaTV)
- [x] 网易cc
- [x] 千度热播
- [x] PandaTV
- [x] 猫耳FM
- [x] Look直播
- [x] WinkTV
- [x] TTingLive(原Flextv)
- [x] PopkonTV
- [x] TwitCasting
- [x] 百度直播
- [x] 微博直播
- [x] 酷狗直播
- [x] TwitchTV
- [x] LiveMe
- [x] 花椒直播
- [x] 流星直播
- [x] ShowRoom
- [x] Acfun
- [x] 映客直播
- [x] 音播直播
- [x] 知乎直播
- [x] CHZZK
- [x] 嗨秀直播
- [x] vv星球直播
- [x] 17Live
- [x] 浪Live
- [x] 畅聊直播
- [x] 飘飘直播
- [x] 六间房直播
- [x] 乐嗨直播
- [x] 花猫直播
- [x] Shopee
- [x] Youtube
- [x] 淘宝
- [x] 京东
- [x] Faceit
- [x] 咪咕
- [x] 连接直播
- [x] 来秀直播
- [x] Picarto
- [ ] 更多平台正在更新中


## 🌱使用说明

- 对于只想使用录制软件的小白用户，进入[Releases](https://github.com/GutsGo/live-recorder/releases) 中下载最新发布的 zip压缩包即可，里面有打包好的录制软件。（有些电脑可能会报毒，直接忽略即可，如果下载时被浏览器屏蔽，请更换浏览器下载）

- 压缩包解压后，在 `config` 文件夹内的 `URL_config.ini` 中添加录制直播间地址，一行一个直播间地址。如果要自定义配置录制，可以修改`config.ini` 文件，推荐将录制格式修改为`ts`。
- 以上步骤都做好后，就可以运行`live-recorder.exe` 程序进行录制了。录制的视频文件保存在同目录下的 `downloads` 文件夹内。

- 另外，如果需要录制TikTok、AfreecaTV等海外平台，请在配置文件中设置开启代理并添加proxy_addr链接 如：`127.0.0.1:7890` （这只是示例地址，具体根据实际填写）。

- 假如`URL_config.ini`文件中添加的直播间地址，有个别直播间暂时不想录制又不想移除链接，可以在对应直播间的链接开头加上`#`，那么将停止该直播间的监测以及录制。

- 软件默认录制清晰度为 `原画` ，如果要单独设置某个直播间的录制画质，可以在添加直播间地址时前面加上画质即可，如`超清，https://live.douyin.com/745964462470` 记得中间要有`,` 分隔。

- 如果要长时间挂着软件循环监测直播，最好循环时间设置长一点（咱也不差没录制到的那几分钟），避免因请求频繁导致被官方封禁IP 。

- 要停止直播录制，Windows平台可执行StopRecording.vbs脚本文件，或者在录制界面使用 `Ctrl+C ` 组合键中断录制，若要停止其中某个直播间的录制，可在`URL_config.ini`文件中的地址前加#，会自动停止对应直播间的录制并正常保存已录制的视频。
- 最后，欢迎右上角给本项目一个star，同时也非常乐意大家提交pr。

&emsp;

直播间链接示例：

```
抖音:
https://live.douyin.com/745964462470
https://v.douyin.com/iQFeBnt/
https://live.douyin.com/yall1102  （链接+抖音号）
https://v.douyin.com/CeiU5cbX  （主播主页地址）

TikTok:
https://www.tiktok.com/@pearlgaga88/live

快手:
https://live.kuaishou.com/u/yall1102

虎牙:
https://www.huya.com/52333

斗鱼:
https://www.douyu.com/3637778?dyshid=
https://www.douyu.com/topic/wzDBLS6?rid=4921614&dyshid=

YY:
https://www.yy.com/22490906/22490906

B站:
https://live.bilibili.com/320

小红书（直播间分享地址):
http://xhslink.com/xpJpfM

bigo直播:
https://www.bigo.tv/cn/716418802

buled直播:
https://app.blued.cn/live?id=Mp6G2R

SOOP:
https://play.sooplive.com/sw7love

网易cc:
https://cc.163.com/583946984

千度热播:
https://qiandurebo.com/web/video.php?roomnumber=33333

PandaTV:
https://www.pandalive.co.kr/live/play/bara0109

猫耳FM:
https://fm.missevan.com/live/868895007

Look直播:
https://look.163.com/live?id=65108820&position=3

WinkTV:
https://www.winktv.co.kr/live/play/anjer1004

FlexTV(TTinglive)::
https://www.flextv.co.kr/channels/593127/live

PopkonTV:
https://www.popkontv.com/live/view?castId=wjfal007&partnerCode=P-00117
https://www.popkontv.com/channel/notices?mcid=wjfal007&mcPartnerCode=P-00117

TwitCasting:
https://twitcasting.tv/c:uonq

百度直播:
https://live.baidu.com/m/media/pclive/pchome/live.html?room_id=9175031377&tab_category

微博直播:
https://weibo.com/l/wblive/p/show/1022:2321325026370190442592

酷狗直播:
https://fanxing2.kugou.com/50428671?refer=2177&sourceFrom=

TwitchTV:
https://www.twitch.tv/gamerbee

LiveMe:
https://www.liveme.com/zh/v/17141543493018047815/index.html

花椒直播:
https://www.huajiao.com/l/345096174

流星直播:
https://www.7u66.com/100960

ShowRoom:
https://www.showroom-live.com/room/profile?room_id=480206  （主播主页地址）

Acfun:
https://live.acfun.cn/live/179922

映客直播:
https://www.inke.cn/liveroom/index.html?uid=22954469&id=1720860391070904

音播直播:
https://live.ybw1666.com/800002949

知乎直播:
https://www.zhihu.com/people/ac3a467005c5d20381a82230101308e9 (主播主页地址)

CHZZK:
https://chzzk.naver.com/live/458f6ec20b034f49e0fc6d03921646d2

嗨秀直播:
https://www.haixiutv.com/6095106

VV星球直播:
https://h5webcdn-pro.vvxqiu.com//activity/videoShare/videoShare.html?h5Server=https://h5p.vvxqiu.com&roomId=LP115924473&platformId=vvstar

17Live:
https://17.live/en/live/6302408

浪Live:
https://www.lang.live/en-US/room/3349463

畅聊直播:
https://live.tlclw.com/106188

飘飘直播:
https://m.pp.weimipopo.com/live/preview.html?uid=91648673&anchorUid=91625862&app=plpl

六间房直播:
https://v.6.cn/634435

乐嗨直播:
https://www.lehaitv.com/8059096

花猫直播:
https://h.catshow168.com/live/preview.html?uid=19066357&anchorUid=18895331

Shopee:
https://sg.shp.ee/GmpXeuf?uid=1006401066&session=802458

Youtube:
https://www.youtube.com/watch?v=cS6zS5hi1w0

淘宝(需cookie):
https://tbzb.taobao.com/live?liveId=532359023188
https://m.tb.cn/h.TWp0HTd

京东:
https://3.cn/28MLBy-E

Faceit:
https://www.faceit.com/zh/players/Compl1/stream

连接直播:
https://show.lailianjie.com/10000258

咪咕直播:
https://www.miguvideo.com/p/live/120000541321

来秀直播:
https://www.imkktv.com/h5/share/video.html?uid=1845195&roomId=1710496

Picarto:
https://www.picarto.tv/cuteavalanche
```

&emsp;

## 🎃源码运行
使用源码运行，可参考下面的步骤。

1.首先拉取或手动下载本仓库项目代码

```bash
git clone https://github.com/GutsGo/live-recorder.git
```

2.进入项目文件夹，安装依赖

```bash
cd live-recorder
```

> [!TIP]
> - 不论你是否已安装 **Python>=3.10** 环境, 都推荐使用 [**uv**](https://github.com/astral-sh/uv) 运行, 因为它可以自动管理虚拟环境和方便地管理 **Python** 版本, **不过这完全是可选的**<br />
> 使用以下命令安装
>    ```bash
>    # 在 macOS 和 Linux 上安装 uv
>    curl -LsSf https://astral.sh/uv/install.sh | sh
>    ```
>    ```powershell
>    # 在 Windows 上安装 uv
>    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
>    ```
> - 如果安装依赖速度太慢, 你可以考虑使用国内 pip 镜像源:<br />
> 在 `pip` 命令使用 `-i` 参数指定, 如 `pip3 install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`<br />
> 或者在 `uv` 命令 `--index` 选项指定, 如 `uv sync --index https://pypi.tuna.tsinghua.edu.cn/simple`

<details>

  <summary>如果已安装 <b>Python>=3.10</b> 环境</summary>

  - :white_check_mark: 在虚拟环境中安装 (推荐)
  
    1. 创建虚拟环境

       - 使用系统已安装的 Python, 不使用 uv
  
         ```bash
         python -m venv .venv
         ```

       - 使用 uv, 默认使用系统 Python, 你可以添加 `--python` 选项指定 Python 版本而不使用系统 Python [uv官方文档](https://docs.astral.sh/uv/concepts/python-versions/)
       
         ```bash
         uv venv
         ```
    
    2. 在终端激活虚拟环境 (在未安装 uv 或你想要手动激活虚拟环境时执行, 若已安装 uv, 可以跳过这一步, uv 会自动激活并使用虚拟环境)
   
       **Bash** 中
       ```bash
       source .venv/Scripts/activate
       ```

       **Powershell** 中
       ```powershell
       .venv\Scripts\activate.ps1
       ```
       
       **Windows CMD** 中
       ```bat
       .venv\Scripts\activate.bat
       ```

    3. 安装依赖
   
       ```bash
       # 使用 pip (若安装太慢或失败, 可使用 `-i` 指定镜像源)
       pip3 install -U pip && pip3 install -r requirements.txt
       # 或者使用 uv (可使用 `--index` 指定镜像源)
       uv sync
       # 或者
       uv pip sync requirements.txt
       ```

  - :x: 在系统 Python 环境中安装 (不推荐)
  
    ```bash
    pip3 install -U pip && pip3 install -r requirements.txt
    ```

</details>

<details>

  <summary>如果未安装 <b>Python>=3.10</b> 环境</summary>

  你可以使用 [**uv**](https://github.com/astral-sh/uv) 安装依赖
   
  ```bash
  # uv 将使用 3.10 及以上的最新 python 发行版自动创建并使用虚拟环境, 可使用 --python 选项指定 python 版本, 参见 https://docs.astral.sh/uv/reference/cli/#uv-sync--python 和 https://docs.astral.sh/uv/reference/cli/#uv-pip-sync--python
  uv sync
  # 或
  uv pip sync requirements.txt
  ```

</details>

3.安装[FFmpeg](https://ffmpeg.org/download.html#build-linux)，如果是Windows系统，这一步可跳过。对于Linux系统，执行以下命令安装

CentOS执行

```bash
yum install epel-release
yum install ffmpeg
```

Ubuntu则执行

```bash
apt update
apt install ffmpeg
```

macOS 执行

**如果已经安装 Homebrew 请跳过这一步**

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

```bash
brew install ffmpeg
```

4.运行程序

```python
python main.py

```
或

```bash
uv run main.py
```

其中Linux系统请使用`python3 main.py` 运行。

&emsp;
## 🐋容器运行

在运行命令之前，请确保您的机器上安装了 [Docker](https://docs.docker.com/get-docker/) 和 [Docker Compose](https://docs.docker.com/compose/install/) 

1.快速启动

最简单方法是运行项目中的 [docker-compose.yaml](https://github.com/GutsGo/live-recorder/blob/main/docker-compose.yaml) 文件，只需简单执行以下命令：

```bash
docker-compose up
```

可选 `-d` 在后台运行。



2.构建镜像(可选)

如果你只想简单的运行程序，则不需要做这一步。Docker镜像仓库中代码版本可能不是最新的，如果要运行本仓库主分支最新代码，可以本地自定义构建，通过修改 [docker-compose.yaml](https://github.com/GutsGo/live-recorder/blob/main/docker-compose.yaml) 文件，如将镜像名修改为 `live-recorder:latest`，并取消 `# build: .` 注释，然后再执行

```bash
docker build -t live-recorder:latest .
docker-compose up
```

或者直接使用下面命令进行构建并启动

```bash
docker-compose -f docker-compose.yaml up
```



3.停止容器实例

```bash
docker-compose stop
```



4.注意事项

①在docker容器内运行本程序之前，请先在配置文件中添加要录制的直播间地址。

②在容器内时，如果手动中断容器运行停止录制，会导致正在录制的视频文件损坏！

**无论哪种运行方式，为避免手动中断或者异常中断导致录制的视频文件损坏的情况，推荐使用 `ts` 格式保存**。

&emsp;

