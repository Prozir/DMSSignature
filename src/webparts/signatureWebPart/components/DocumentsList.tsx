import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, ConstrainMode, IColumn, SelectionMode, Spinner, SpinnerSize, Dropdown, IDropdownOption } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './SignatureWebPart.module.scss';
import type { DocumentDisplayMode } from './ISignatureWebPartProps';

type StatusFilter = 'Pending' | 'Signed' | 'Rejected' | 'Total';

const statusFilterOptions: IDropdownOption[] = [
  { key: 'Pending', text: 'Pending' },
  { key: 'Signed', text: 'Signed' },
  { key: 'Rejected', text: 'Rejected' },
  { key: 'Total', text: 'Total' }
];
export interface IDocumentListItem {
  id: number;
  title?: string;
  documentId?: string;
  documentName?: string;
  documentNumber?: string;
  trader?: string;
  accountCode?: string;
  approverEmail?: string;
  approvalStatus?: string;
  comments?: string;
  isTaskActive?: boolean;
  created?: string;
  modified?: string;
  attachmentFileName?: string;
  attachmentServerRelativeUrl?: string;
}

export interface IDocumentsListProps {
  items: IDocumentListItem[];
  isLoading: boolean;
  displayMode: DocumentDisplayMode;
  onOpenItem: (item: IDocumentListItem) => void;
  onShowHistory: (item: IDocumentListItem) => void;
}

const getDateTimeValue = (value?: string): string => value ? new Date(value).toLocaleString() : '';

const isPending = (item: IDocumentListItem): boolean =>
  (item.approvalStatus || '').toLowerCase().indexOf('pending') !== -1;

const renderActionIcon = (item: IDocumentListItem, onOpenItem: (item: IDocumentListItem) => void, className?: string): React.ReactElement => {
  const isPendingItem: boolean = isPending(item);

  return (
    <Icon
      iconName="InsertSignatureLine"
      className={className}
      title={isPendingItem ? `Open ${item.documentName || item.title || 'document'}` : 'Available only for Pending items'}
      aria-label={isPendingItem ? `Open ${item.documentName || item.title || 'document'}` : 'Action disabled for non-pending status'}
      aria-disabled={!isPendingItem}
      role={isPendingItem ? 'button' : undefined}
      tabIndex={isPendingItem ? 0 : -1}
      onClick={isPendingItem ? (() => onOpenItem(item)) : undefined}
      onKeyDown={(e) => {
        if (!isPendingItem) {
          return;
        }

        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          onOpenItem(item);
        }
      }}
      styles={{
        root: {
          cursor: isPendingItem ? 'pointer' : 'not-allowed',
          fontSize: 16,
          color: isPendingItem ? 'var(--themePrimary)' : '#94a3b8',
          selectors: {
            ':hover': { color: isPendingItem ? 'var(--themeDarkAlt)' : '#94a3b8' },
          },
        },
      }}
    />
  );
};

const renderHistoryIcon = (item: IDocumentListItem, onShowHistory: (item: IDocumentListItem) => void, className?: string): React.ReactElement => (
  <Icon
    iconName="History"
    className={className}
    title={`View approval history for ${item.documentName || item.title || 'document'}`}
    aria-label={`View approval history for ${item.documentName || item.title || 'document'}`}
    role="button"
    tabIndex={0}
    onClick={() => onShowHistory(item)}
    onKeyDown={(event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        onShowHistory(item);
      }
    }}
    styles={{
      root: {
        cursor: 'pointer',
        fontSize: 16,
        color: 'var(--themePrimary)',
        selectors: {
          ':hover': { color: 'var(--themeDarkAlt)' }
        }
      }
    }}
  />
);

const getColumns = (onOpenItem: (item: IDocumentListItem) => void, onShowHistory: (item: IDocumentListItem) => void): IColumn[] => [
  {
    key: 'columnTaskId',
    name: 'Task ID',
    fieldName: 'id',
    minWidth: 70,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.id
  },
  {
    key: 'columnDocumentName',
    name: 'Document Name',
    fieldName: 'documentName',
    minWidth: 150,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.documentName || item.title || ''
  },
  {
    key: 'columnDocumentNumber',
    name: 'Document Number',
    fieldName: 'documentNumber',
    minWidth: 120,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.documentNumber || ''
  },
  {
    key: 'columnTrader',
    name: 'Trader',
    fieldName: 'trader',
    minWidth: 120,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.trader || ''
  },
  {
    key: 'columnAccountCode',
    name: 'Account Code',
    fieldName: 'accountCode',
    minWidth: 120,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.accountCode || ''
  },
  {
    key: 'columnApprovalStatus',
    name: 'Approval Status',
    fieldName: 'approvalStatus',
    minWidth: 140,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.approvalStatus || ''
  },
  {
    key: 'columnCreated',
    name: 'Created',
    fieldName: 'created',
    minWidth: 140,
    isResizable: true,
    onRender: (item: IDocumentListItem) => getDateTimeValue(item.created)
  },
  {
    key: 'columnModified',
    name: 'Modified',
    fieldName: 'modified',
    minWidth: 140,
    isResizable: true,
    onRender: (item: IDocumentListItem) => getDateTimeValue(item.modified)
  },
  {
    key: 'columnAction',
    name: 'Action',
    fieldName: 'action',
    minWidth: 80,
    isResizable: false,
    onRender: (item: IDocumentListItem) => (
      <div className={styles.documentListActions}>
        {renderActionIcon(item, onOpenItem)}
        {renderHistoryIcon(item, onShowHistory)}
      </div>
    )
  }
];

export default function DocumentsList(props: IDocumentsListProps): React.ReactElement {
  const [selectedStatus, setSelectedStatus] = React.useState<StatusFilter>('Pending');

  const filteredItems: IDocumentListItem[] = React.useMemo(() => {
    const normalizedStatus: string = selectedStatus.toLowerCase();
    const statusFiltered: IDocumentListItem[] = selectedStatus === 'Total'
      ? props.items
      : props.items.filter((item: IDocumentListItem) =>
          (item.approvalStatus || '').toLowerCase().indexOf(normalizedStatus) !== -1
        );

    return statusFiltered.slice().sort((a: IDocumentListItem, b: IDocumentListItem) => b.id - a.id);
  }, [props.items, selectedStatus]);

  if (props.isLoading) {
    return <Spinner label="Loading tasks..." size={SpinnerSize.medium} />;
  }

  return (
    <div>
      <div style={{ marginBottom: 12, maxWidth: 220 }}>
        <Dropdown
          label="Filter by status"
          selectedKey={selectedStatus}
          options={statusFilterOptions}
          onChange={(_, option) => setSelectedStatus((option?.key as StatusFilter) || 'Pending')}
        />
      </div>

      {filteredItems.length === 0 ? (
        <div>No items found for {selectedStatus} status.</div>
      ) : props.displayMode === 'cards' ? (
        <div className={styles.documentsCards}>
          {filteredItems.map((item: IDocumentListItem) => (
            <div key={item.id} className={styles.documentCard}>
              <div className={styles.documentCardHeader}>
                <span className={styles.documentCardTitle}>{item.documentName || item.title || 'Untitled document'}</span>
                <div className={styles.documentCardAction}>
                  <div className={styles.documentListActions}>
                    {renderActionIcon(item, props.onOpenItem, styles.documentActionIcon)}
                  </div>
                </div>
              </div>

              <div className={styles.documentCardBody}>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Task ID</span>
                  <span className={styles.documentCardValue}>{item.id}</span>
                </div>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Document Number</span>
                  <span className={styles.documentCardValue}>{item.documentNumber || '-'}</span>
                </div>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Trader</span>
                  <span className={styles.documentCardValue}>{item.trader || '-'}</span>
                </div>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Account Code</span>
                  <span className={styles.documentCardValue}>{item.accountCode || '-'}</span>
                </div>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Approval Status</span>
                  <span className={`${styles.documentCardValue} ${styles.documentCardStatusValue}`}>
                    <span>{item.approvalStatus || '-'}</span>
                    {renderHistoryIcon(item, props.onShowHistory, styles.documentActionIcon)}
                  </span>
                </div>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Created</span>
                  <span className={styles.documentCardValue}>{getDateTimeValue(item.created) || '-'}</span>
                </div>
                <div className={styles.documentCardField}>
                  <span className={styles.documentCardLabel}>Modified</span>
                  <span className={styles.documentCardValue}>{getDateTimeValue(item.modified) || '-'}</span>
                </div>
              </div>

            </div>
          ))}
        </div>
      ) : (
        <DetailsList
          items={filteredItems}
          columns={getColumns(props.onOpenItem, props.onShowHistory)}
          setKey="documentsList"
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          constrainMode={ConstrainMode.unconstrained}
          compact={true}
          //ariaLabel="Signed documents list"
        />
      )}
    </div>
  );
}
