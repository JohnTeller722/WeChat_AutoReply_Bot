# WeChat Auto Reply Bot

一个基于 Python 的微信自动回复机器人，支持规则匹配和 AI 对话。

## 功能特性

1. 自动监控并回复微信消息
2. 支持规则匹配和 AI 对话双重回复机制
3. 可配置多个联系人的自动回复
4. 支持自定义回复前缀
5. 具有消息频率限制功能
6. 保持对话上下文记忆
7. 支持定时重置统计数据

## 下一步开发计划

1. 支持RAG检索
2. 支持表情包识别与图片理解能力
3. 微调大模型，使得回答语气风格更加个性化
4. 微调大模型，使得大模型回答更像日常聊天
5. 增强基于规则库的问答，一方面扩充规则数量，一方面进行模糊匹配检索
6. 支持并发处理
7. 支持接入群聊（能够识别不同人）
8. 扩充角色，不仅可以是自动回复机器人，还可以当客服或者猫娘等自定义角色，或者当自动总结与分析聊天记录的秘书

## 安装说明

1.安装依赖

```bash
pip install pandas uiautomation pyyaml openai
```

2.准备配置文件

创建 config.yaml：配置全局设置和联系人列表

创建 rules.csv：配置关键词回复规则

## 配置说明

### config.yaml 配置项

```yaml
global:
  max_replies: 20          # 每个联系人最大回复次数
  rules_file: "rules.csv"  # 规则文件路径
  default_reply: "..."     # 默认回复语
  reset_interval: 60       # 重置间隔（分钟）
  my_wx_name: "xxx"       # 主窗口联系人名称

contacts:                  # 监听的联系人列表
  - name: "联系人名称"
    enabled: true         # 是否启用
    prefix: "前缀消息"    # 首次回复的前缀
```

### rules.csv 格式

```csv
序号,关键词,回复内容
1,关键词1,回复内容1
2,关键词2,回复内容2
```

## 使用方法

1. 配置好 config.yaml 和 rules.csv

2. 打开微信电脑版并登录

3. 运行程序

```bash
python autoReply.py
```

## 实现原理

1. 消息监控

- 使用 uiautomation 监控微信窗口
- 检测新消息提醒
- 自动切换会话窗口

2. 回复机制

- 优先匹配 rules.csv 中的规则
- 如无匹配规则，使用 AI 模型生成回复
- 保持对话上下文提升回复质量

3. 限制机制

- 记录每个联系人的回复次数
- 达到上限后自动停止回复
- 定时重置统计数据

4. 上下文管理

- 为每个联系人维护独立的对话历史
- 记录最近几轮对话内容
- 保存到独立的 JSON 文件
