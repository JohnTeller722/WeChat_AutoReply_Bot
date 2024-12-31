import pandas as pd
from uiautomation import WindowControl
import time
import yaml
import json
import os
from datetime import datetime
from openai import OpenAI
from collections import deque

def load_config():
    """加载配置文件"""
    try:
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            
            # 处理全局配置
            global_config = config.get('global', {})
            
            # 处理联系人列表
            contacts = []
            for contact in config.get('contacts', []):
                if contact.get('enabled', True):
                    contacts.append({
                        'name': contact['name'],
                        'prefix': contact.get('prefix', 
                            "我是自动回复机器人张和，我的主人现在失踪了，我先代为回复~")
                    })
            
            return {
                'MAX_REPLIES': global_config.get('max_replies', 20),
                'DATA_FILE': global_config.get('rules_file', 'rules.csv'),
                'DEFAULT_REPLY': global_config.get('default_reply', 
                    '自动回复机器人出故障了www，请联系我的master修一下QwQ'),
                'RESET_INTERVAL': global_config.get('reset_interval', 60),  # 默认60分钟
                'CONTACTS': contacts,
                'MY_WX_NAME': global_config.get('my_wx_name', 'John Teller')
            }
    except Exception as e:
        print(f"加载配置文件失败: {str(e)}")
        return {
            'MAX_REPLIES': 20,
            'DATA_FILE': 'rules.csv',
            'DEFAULT_REPLY': '自动回复机器人出故障了www，请联系我的master修一下QwQ',
            'RESET_INTERVAL': 60,
            'CONTACTS': [
                {
                    'name': 'John Teller',
                    'prefix': "我是自动回复机器人张和，我的主人现在失踪了，我先代为回复~"
                }
            ],
            'MY_WX_NAME': 'John Teller'
        }

# 加载配置
CONFIG = load_config()

class WeChatBot:
    def __init__(self):
        self.wx = None
        self.df = None
        self.config = load_config()
        self.MAX_REPLIES = self.config['MAX_REPLIES']
        self.RESET_INTERVAL = self.config.get('RESET_INTERVAL', 60)  # 默认60分钟
        self.prefix_sent = set()
        self.context_dir = self._create_context_dir()
        self.contexts = {}  # 存储每个联系人的上下文
        
        # 初始化时重置统计
        self._reset_reply_stats()
        self.reply_stats = self._load_reply_stats()

    def _create_context_dir(self):
        """创建会话上下文目录"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        context_dir = f'context_{timestamp}'
        os.makedirs(context_dir, exist_ok=True)
        print(f"创建会话上下文目录: {context_dir}")
        return context_dir

    def _reset_reply_stats(self):
        """重置回复统计"""
        try:
            stats = {
                'last_reset': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'contacts': {}
            }
            with open('reply_stats.json', 'w', encoding='utf-8') as f:
                json.dump(stats, f, ensure_ascii=False, indent=2)
            print("已重置回复统计")
        except Exception as e:
            print(f"重置回复统计失败: {str(e)}")
    
    def _load_reply_stats(self):
        """加载回复统计数据"""
        try:
            if os.path.exists('reply_stats.json'):
                with open('reply_stats.json', 'r', encoding='utf-8') as f:
                    stats = json.load(f)
                    
                    # 检查是否需要重置
                    if self._should_reset_stats(stats.get('last_reset')):
                        self._reset_reply_stats()
                        return self._load_reply_stats()
                    
                    return stats.get('contacts', {})
            return {}
        except Exception as e:
            print(f"加载回复统计失败: {str(e)}")
            return {}
    
    def _should_reset_stats(self, last_reset_str):
        """检查是否应该重置统计"""
        try:
            if not last_reset_str:
                return True
                
            last_reset = datetime.strptime(last_reset_str, '%Y-%m-%d %H:%M:%S')
            time_diff = datetime.now() - last_reset
            return time_diff.total_seconds() / 60 >= self.RESET_INTERVAL
        except Exception as e:
            print(f"检查是否应该重置统计失败: {str(e)}")
            return True
    
    def _save_reply_stats(self):
        """保存回复统计数据"""
        try:
            stats = {
                'last_reset': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'contacts': self.reply_stats
            }
            with open('reply_stats.json', 'w', encoding='utf-8') as f:
                json.dump(stats, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存回复统计失败: {str(e)}")
    
    def _check_reply_limit(self, contact):
        """
        检查是否超出回复限制
        返回: (是否超限, 是否需要发送通知)
        """
        stats = self.reply_stats.get(contact, {'count': 0, 'limit_notified': False})
        is_limit_reached = stats['count'] >= self.MAX_REPLIES
        need_notification = is_limit_reached and not stats.get('limit_notified', False)
        return is_limit_reached, need_notification
    
    def _update_reply_stats(self, contact):
        """更新回复统计"""
        if contact not in self.reply_stats:
            self.reply_stats[contact] = {
                'count': 0,
                'first_reply': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
        self.reply_stats[contact]['count'] += 1
        self.reply_stats[contact]['last_reply'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self._save_reply_stats()
    
    def _send_limit_notification(self, contact):
        """发送达到限制的通知"""
        try:
            message = f"已达到自动回复上限({self.MAX_REPLIES}条)，请等待我的主人回来后联系您~"
            self.wx.SendKeys(message, waitTime=0)
            self.wx.SendKeys('{Enter}', waitTime=1)
            
            # 更新通知状态
            if contact not in self.reply_stats:
                self.reply_stats[contact] = {}
            self.reply_stats[contact]['limit_notified'] = True
            self._save_reply_stats()
            
            print(f"已发送限制通知给 {contact}")
        except Exception as e:
            print(f"发送限制通知失败: {str(e)}")
    
    def _send_message(self, message, contact):
        """发送消息"""
        contact_name = contact['name']
        
        # 检查是否需要发送前缀
        if contact_name not in self.prefix_sent:
            prefix = contact['prefix']
            self.wx.SendKeys(prefix, waitTime=0)
            self.wx.SendKeys('{Enter}', waitTime=1)
            self.prefix_sent.add(contact_name)
            
        # 发送实际消息
        self.wx.SendKeys(message, waitTime=0)
        self.wx.SendKeys('{Enter}', waitTime=1)
        print(f"已回复 {contact_name}: {message}")
        
        # 更新统计
        self._update_reply_stats(contact_name)
        return True
    
    def setup(self):
        """初始化微信窗口和数据"""
        try:
            # 初始化微信窗口
            self.wx = WindowControl(Name='微信', searchDepth=1)
            self.wx.ListControl()
            self.wx.SwitchToThisWindow()
            
            # 加载回复数据
            self.df = pd.read_csv(CONFIG['DATA_FILE'], encoding='utf-8')
            print("数据加载成功：\n", self.df)
            
            return True
        except Exception as e:
            print(f"设置失败: {str(e)}")
            return False
    
    def chat(self, message, contact_name):
        """聊天"""
        # 优先匹配规则
        reply = self._match_and_reply(message)
        if reply != CONFIG['DEFAULT_REPLY']:
            self._update_context(contact_name, message, reply)
            return reply
        
        # 如果规则中没有匹配到，则使用AI回复
        return self.ai_reply(message, contact_name)

    def _match_and_reply(self, message):
        """匹配消息并生成回复"""
        # 在数据中查找匹配的回复
        matches = self.df.apply(
            lambda x: x['回复内容'] if x['关键词'] in message else None, 
            axis=1
        ).dropna()
        
        if not matches.empty:
            reply = matches.iloc[0]
            return reply.replace('{br}', '{Shift}{Enter}')
        return CONFIG['DEFAULT_REPLY']
    
    def ai_reply(self, message, contact_name):
        """处理聊天消息"""
        # 加载上下文
        context = self._load_context(contact_name)
        
        # 构建消息列表
        messages = [
            {"role": "system", "content": "你叫\"张和\"，你的主人叫\"JohnTeller\"，他现在失踪了，你现在是在代替他和别人聊天，尽量表现得自然，就像人类自然聊天一样，注意语气友好，并且尽量回复简单，严禁使用表情包，如果要使用请使用颜文字代替"}
        ]
        
        # 添加历史上下文
        for ctx in context:
            messages.extend([
                {"role": "user", "content": ctx['user']},
                {"role": "assistant", "content": ctx['bot']}
            ])
        
        # 添加当前消息
        messages.append({"role": "user", "content": message})
        
        client = OpenAI(
            api_key="8fd964e9848dc636803bc68aaa51b0a6.0asVfre3nxkWIjvj",
            base_url="https://open.bigmodel.cn/api/paas/v4/"
        )
        try:
            # 调用API获取回复
            response = client.chat.completions.create(
                model="glm-4-airx",
                messages=messages
            )
            reply = response.choices[0].message.content
            
            # 更新上下文
            self._update_context(contact_name, message, reply)
            
            return reply
        except Exception as e:
            print(f"获取回复失败: {str(e)}")
            return self.config['DEFAULT_REPLY']
    
    def _switch_to_main_contact(self):
        """切换到主要联系人窗口"""
        try:
            conversations = self.wx.ListControl(Name='会话').GetChildren()
            for conv in conversations:
                if CONFIG['MY_WX_NAME'] in conv.Name:  # 可以后续改为配置项
                    conv.Click(simulateMove=False)
                    time.sleep(0.5)  # 等待窗口切换
                    print(f"已切换至{CONFIG['MY_WX_NAME']}")
                    return True
            print(f"未找到{CONFIG['MY_WX_NAME']}")
            return False
        except Exception as e:
            print(f"切换到主要联系人窗口失败: {str(e)}")
            return False

    def process_new_messages(self):
        """处理新消息"""
        try:
            conversations = self.wx.ListControl(Name='会话').GetChildren()
            
            for conv in conversations:
                current_contact = None
                for contact in self.config['CONTACTS']:
                    if contact['name'] in conv.Name:
                        current_contact = contact
                        break
                
                if not current_contact:
                    continue
                
                new_msg_count = self._has_unread_messages(conv)
                if new_msg_count == 0:
                    continue
                
                print(f"正在处理 {current_contact['name']} 的新消息")
                
                conv.Click(simulateMove=False)
                time.sleep(0.5)
                
                all_messages = self.wx.ListControl(Name='消息').GetChildren()
                message_list = all_messages[-new_msg_count:]
                print(f"获取到{len(message_list)}条新消息")
                
                for msg in reversed(message_list):
                    is_limit_reached, need_notification = self._check_reply_limit(current_contact['name'])
                    if is_limit_reached:
                        if need_notification:
                            self._send_limit_notification(current_contact['name'])
                        print(f"{current_contact['name']} 已达到回复上限({self.MAX_REPLIES}条)，停止回复")
                        break
                    
                    reply = self.chat(msg.Name, current_contact['name'])
                    if not self._send_message(reply, current_contact):
                        break
                
                print(f"完成处理 {current_contact['name']} 的新消息")
                
                if current_contact['name'] != CONFIG['MY_WX_NAME']:
                    time.sleep(0.5)
                    self._switch_to_main_contact()
            
        except Exception as e:
            print(f"处理消息时出错: {str(e)}")
            self._switch_to_main_contact()
    
    def _has_unread_messages(self, conversation):
        """
        检查会话是否有未读消息
        返回: 未读消息数量，如果没有则返回0
        """
        try:
            conversation_name = conversation.Name
            if "条新消息" in conversation_name:
                # 找到"条新消息"的位置
                marker_index = conversation_name.find("条新消息")
                # 从后往前找数字
                count = ""
                index = marker_index - 1
                while index >= 0 and conversation_name[index].isdigit():
                    count = conversation_name[index] + count
                    index -= 1
                return int(count) if count else 0
            return 0
        except Exception as e:
            print(f"检查未读消息状态时出错: {str(e)}")
            return 0

    def _get_context_file(self, contact_name):
        """获取联系人的上下文文件路径"""
        return os.path.join(self.context_dir, f'{contact_name}_context.json')
    
    def _load_context(self, contact_name):
        """加载联系人的对话上下文"""
        if contact_name not in self.contexts:
            context_file = self._get_context_file(contact_name)
            if os.path.exists(context_file):
                try:
                    with open(context_file, 'r', encoding='utf-8') as f:
                        self.contexts[contact_name] = deque(json.load(f), maxlen=5)
                except Exception as e:
                    print(f"加载上下文失败: {str(e)}")
                    self.contexts[contact_name] = deque(maxlen=5)
            else:
                self.contexts[contact_name] = deque(maxlen=5)
        return self.contexts[contact_name]
    
    def _save_context(self, contact_name):
        """保存联系人的对话上下文"""
        if contact_name in self.contexts:
            context_file = self._get_context_file(contact_name)
            try:
                with open(context_file, 'w', encoding='utf-8') as f:
                    json.dump(list(self.contexts[contact_name]), f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"保存上下文失败: {str(e)}")
    
    def _update_context(self, contact_name, user_message, bot_reply):
        """更新对话上下文"""
        context = self._load_context(contact_name)
        context.append({
            'user': user_message,
            'bot': bot_reply,
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })
        self._save_context(contact_name)

    def run(self):
        """运行主循环"""
        if not self.setup():
            return
            
        print("开始监听新消息...")
        last_check_time = datetime.now()
        
        while True:
            try:
                # 检查是否需要重置统计
                now = datetime.now()
                if (now - last_check_time).total_seconds() / 60 >= self.RESET_INTERVAL:
                    self._reset_reply_stats()
                    self.reply_stats = self._load_reply_stats()
                    last_check_time = now
                
                self.process_new_messages()
                time.sleep(1)
            except Exception as e:
                print(f"处理消息时出错: {str(e)}")
                time.sleep(1)

if __name__ == "__main__":
    bot = WeChatBot()
    bot.run()