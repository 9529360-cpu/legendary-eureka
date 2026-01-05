import React, { useEffect, useRef, useState } from 'react';
import {
  FluentProvider,
  webLightTheme,
  Button,
  Textarea,
  Avatar,
  makeStyles,
  shorthands,
  tokens,
  Caption1,
  Subtitle1,
  Body1,
} from '@fluentui/react-components';
import { SendRegular, QuestionCircleRegular } from '@fluentui/react-icons';

// 1. 使用 Fluent UI 标准方案定义样式，解决 CSS 渲染冲突
const useStyles = makeStyles({
  container: {
    height: '100vh',
    display: 'flex',
    flexDirection: 'column',
    backgroundColor: tokens.colorNeutralBackground2,
  },
  header: {
    padding: '12px 16px',
    display: 'flex',
    alignItems: 'center',
    ...shorthands.gap('12px'),
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  messageList: {
    flexGrow: 1,
    overflowY: 'auto',
    padding: '16px',
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('16px'),
  },
  msgRow: {
    display: 'flex',
    width: '100%',
  },
  msgUser: { justifyContent: 'flex-end' },
  msgAI: { justifyContent: 'flex-start' },
  bubbleUser: {
    maxWidth: '85%',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    padding: '10px 14px',
    borderRadius: '12px 12px 2px 12px',
    boxShadow: tokens.shadow4,
  },
  bubbleAI: {
    maxWidth: '85%',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    padding: '10px 14px',
    borderRadius: '2px 12px 12px 12px',
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    boxShadow: tokens.shadow2,
  },
  inputArea: {
    padding: '16px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    display: 'flex',
    alignItems: 'flex-end',
    ...shorthands.gap('8px'),
  },
  payloadCard: {
    marginTop: '8px',
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke3),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
  },
  headerRight: {
    flexGrow: 1,
  },
  bubbleText: {
    whiteSpace: 'pre-wrap',
  },
  payloadPre: {
    fontSize: '10px',
    margin: 0,
  },
  inputTextarea: {
    flexGrow: 1,
  }
});

type Message = { id: string; role: 'assistant' | 'user'; text: string; payload?: any };

export default function App() {
  const styles = useStyles();
  const [messages, setMessages] = useState<Message[]>([
    { id: '1', role: 'assistant', text: '你好！我是 Excel Copilot。请问今天需要我帮您分析数据还是检查表格？' }
  ]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const scrollRef = useRef<HTMLDivElement>(null);

  // 自动滚动到底部
  useEffect(() => {
    if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
  }, [messages]);

  const handleSend = async () => {
    if (!input.trim() || isLoading) return;
    
    const userMsg: Message = { id: Date.now().toString(), role: 'user', text: input };
    setMessages(prev => [...prev, userMsg]);
    setInput('');
    setIsLoading(true);

    // TODO: 这里将来替换为你已写好的 DeepSeek 后端请求逻辑
    setTimeout(() => {
      const aiMsg: Message = { 
        id: (Date.now()+1).toString(), 
        role: 'assistant', 
        text: `已收到指令 "${userMsg.text}"。我正在调取 Excel 接口进行处理...` 
      };
      setMessages(prev => [...prev, aiMsg]);
      setIsLoading(false);
    }, 1000);
  };

  return (
    <FluentProvider theme={webLightTheme}>
      <div className={styles.container}>
        {/* 顶部状态栏 */}
        <header className={styles.header}>
          <Avatar color="brand" name="Excel Copilot" initials="EC" />
          <div className={styles.headerRight}>
            <Subtitle1 block>Excel Copilot</Subtitle1>
            <Caption1 italic color="neutralTertiary">AI 驱动 · 智能审计</Caption1>
          </div>
          <Button icon={<QuestionCircleRegular />} appearance="subtle" />
        </header>

        {/* 聊天消息区域 */}
        <main className={styles.messageList} ref={scrollRef}>
          {messages.map((m) => (
            <div key={m.id} className={`${styles.msgRow} ${m.role === 'user' ? styles.msgUser : styles.msgAI}`}>
              <div className={m.role === 'user' ? styles.bubbleUser : styles.bubbleAI}>
                <Body1 block className={styles.bubbleText}>{m.text}</Body1>
                {m.payload && (
                  <div className={styles.payloadCard}>
                    <Caption1 strong block>感知数据结果：</Caption1>
                    <pre className={styles.payloadPre}>{JSON.stringify(m.payload, null, 2)}</pre>
                  </div>
                )}
              </div>
            </div>
          ))}
          {isLoading && <Caption1 className={styles.msgAI}>Copilot 正在思考...</Caption1>}
        </main>

        {/* 输入区域 */}
        <footer className={styles.inputArea}>
          <Textarea
            value={input}
            onChange={(_, d) => setInput(d.value)}
            onKeyDown={(e) => e.key === 'Enter' && !e.shiftKey && (e.preventDefault(), handleSend())}
            placeholder="请输入指令..."
            resize="none"
            rows={1}
            className={styles.inputTextarea}
          />
          <Button 
            appearance="primary" 
            icon={<SendRegular />} 
            onClick={handleSend}
            disabled={isLoading}
          />
        </footer>
      </div>
    </FluentProvider>
  );
}