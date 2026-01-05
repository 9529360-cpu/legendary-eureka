import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Title2,
  Body2,
  Caption1,
  Button,
  Input,
} from "@fluentui/react-components";
import {
  HistoryRegular,
  PanelLeftContractRegular,
  SearchRegular,
  DeleteRegular,
  ChatRegular,
} from "@fluentui/react-icons";

export interface HistoryItemData {
  id: string;
  title: string;
  preview: string;
  timestamp: Date;
  messageCount: number;
}

interface HistorySidebarProps {
  historyItems: HistoryItemData[];
  selectedId?: string;
  onSelectItem: (id: string) => void;
  onDeleteItem: (id: string) => void;
  onNewChat: () => void;
  onCollapse: () => void;
}

const useStyles = makeStyles({
  sidebar: {
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  header: {
    ...shorthands.padding("16px"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  headerTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  searchContainer: {
    ...shorthands.padding("12px"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  content: {
    flex: 1,
    overflow: "auto",
    scrollbarGutter: "stable",
    ...shorthands.padding("12px"),
  },
  historyItem: {
    ...shorthands.padding("10px", "12px"),
    ...shorthands.borderRadius("6px"),
    cursor: "pointer",
    marginBottom: "4px",
    transition: "background-color 0.15s ease",
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  historyItemSelected: {
    backgroundColor: tokens.colorNeutralBackground3,
    borderLeft: `3px solid ${tokens.colorBrandBackground}`,
  },
  historyItemHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  historyItemTitle: {
    fontWeight: "600",
    flex: 1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  historyItemPreview: {
    color: tokens.colorNeutralForeground4,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  historyItemMeta: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  deleteButton: {
    opacity: 0,
    transition: "opacity 0.15s ease",
    ":hover": {
      opacity: 1,
    },
  },
  newChatButton: {
    marginTop: "auto",
    ...shorthands.padding("12px"),
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  emptyState: {
    textAlign: "center",
    ...shorthands.padding("24px"),
    color: tokens.colorNeutralForeground4,
  },
});

const HistorySidebar: React.FC<HistorySidebarProps> = ({
  historyItems,
  selectedId,
  onSelectItem,
  onDeleteItem,
  onNewChat,
  onCollapse,
}) => {
  const styles = useStyles();
  const [searchQuery, setSearchQuery] = React.useState("");
  const [hoveredId, setHoveredId] = React.useState<string | null>(null);

  const filteredItems = React.useMemo(() => {
    if (!searchQuery) return historyItems;
    const query = searchQuery.toLowerCase();
    return historyItems.filter(
      (item) =>
        item.title.toLowerCase().includes(query) ||
        item.preview.toLowerCase().includes(query)
    );
  }, [historyItems, searchQuery]);

  const formatDate = (date: Date): string => {
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));

    if (days === 0) {
      return date.toLocaleTimeString("zh-CN", {
        hour: "2-digit",
        minute: "2-digit",
      });
    } else if (days === 1) {
      return "昨天";
    } else if (days < 7) {
      return `${days} 天前`;
    } else {
      return date.toLocaleDateString("zh-CN", {
        month: "short",
        day: "numeric",
      });
    }
  };

  return (
    <div className={styles.sidebar}>
      <div className={styles.header}>
        <div className={styles.headerTitle}>
          <HistoryRegular />
          <Title2>历史记录</Title2>
        </div>
        <Button
          appearance="subtle"
          size="small"
          icon={<PanelLeftContractRegular />}
          onClick={onCollapse}
        />
      </div>

      <div className={styles.searchContainer}>
        <Input
          placeholder="搜索历史..."
          contentBefore={<SearchRegular />}
          value={searchQuery}
          onChange={(e, data) => setSearchQuery(data.value)}
        />
      </div>

      <div className={styles.content}>
        {filteredItems.length === 0 ? (
          <div className={styles.emptyState}>
            <ChatRegular style={{ fontSize: 32, marginBottom: 8 }} />
            <Body2>暂无历史记录</Body2>
            <Caption1>开始新对话后将在此显示</Caption1>
          </div>
        ) : (
          filteredItems.map((item) => (
            <div
              key={item.id}
              className={`${styles.historyItem} ${
                selectedId === item.id ? styles.historyItemSelected : ""
              }`}
              onClick={() => onSelectItem(item.id)}
              onMouseEnter={() => setHoveredId(item.id)}
              onMouseLeave={() => setHoveredId(null)}
            >
              <div className={styles.historyItemHeader}>
                <Body2 className={styles.historyItemTitle}>{item.title}</Body2>
                {hoveredId === item.id && (
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<DeleteRegular />}
                    className={styles.deleteButton}
                    onClick={(e) => {
                      e.stopPropagation();
                      onDeleteItem(item.id);
                    }}
                  />
                )}
              </div>
              <Caption1 className={styles.historyItemPreview}>
                {item.preview}
              </Caption1>
              <div className={styles.historyItemMeta}>
                <Caption1>{formatDate(item.timestamp)}</Caption1>
                <Caption1>{item.messageCount} 条消息</Caption1>
              </div>
            </div>
          ))
        )}
      </div>

      <div className={styles.newChatButton}>
        <Button appearance="primary" style={{ width: "100%" }} onClick={onNewChat}>
          新建对话
        </Button>
      </div>
    </div>
  );
};

export default HistorySidebar;
