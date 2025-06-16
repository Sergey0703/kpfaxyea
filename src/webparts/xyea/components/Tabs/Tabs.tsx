// src/webparts/xyea/components/Tabs/Tabs.tsx

import * as React from 'react';
import styles from './Tabs.module.scss';

export interface ITabItem {
  key: string;
  label: string;
  content: React.ReactNode;
}

export interface ITabsProps {
  items: ITabItem[];
  defaultActiveKey?: string;
  onTabChange?: (activeKey: string) => void;
}

export interface ITabsState {
  activeKey: string;
}

export default class Tabs extends React.Component<ITabsProps, ITabsState> {
  
  constructor(props: ITabsProps) {
    super(props);
    
    this.state = {
      activeKey: props.defaultActiveKey || (props.items.length > 0 ? props.items[0].key : '')
    };
  }

  private handleTabClick = (key: string): void => {
    this.setState({ activeKey: key });
    if (this.props.onTabChange) {
      this.props.onTabChange(key);
    }
  }

  public render(): React.ReactElement<ITabsProps> {
    const { items } = this.props;
    const { activeKey } = this.state;

    const activeItem = items.find(item => item.key === activeKey);

    return (
      <div className={styles.tabs}>
        <div className={styles.tabsHeader}>
          {items.map((item) => (
            <button
              key={item.key}
              className={`${styles.tabButton} ${activeKey === item.key ? styles.active : ''}`}
              onClick={() => this.handleTabClick(item.key)}
            >
              {item.label}
            </button>
          ))}
        </div>
        <div className={styles.tabsContent}>
          {activeItem && activeItem.content}
        </div>
      </div>
    );
  }
}