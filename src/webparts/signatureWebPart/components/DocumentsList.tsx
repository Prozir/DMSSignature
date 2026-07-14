import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, ConstrainMode, IColumn, SelectionMode, Spinner, SpinnerSize, Dropdown, IDropdownOption } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';

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
  onOpenItem: (item: IDocumentListItem) => void;
}

const getColumns = (onOpenItem: (item: IDocumentListItem) => void): IColumn[] => [
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
    onRender: (item: IDocumentListItem) => item.created ? new Date(item.created).toLocaleString() : ''
  },
  {
    key: 'columnModified',
    name: 'Modified',
    fieldName: 'modified',
    minWidth: 140,
    isResizable: true,
    onRender: (item: IDocumentListItem) => item.modified ? new Date(item.modified).toLocaleString() : ''
  },
  {
    key: 'columnAction',
    name: 'Action',
    fieldName: 'action',
    minWidth: 80,
    isResizable: false,
    onRender: (item: IDocumentListItem) => {
      const isPendingItem = (item.approvalStatus || '').toLowerCase().indexOf('pending') !== -1;

      return (
      /*<IconButton
        iconProps={{ iconName: 'InsertSignatureLine' }}
        title={`Open ${item.documentName || item.title || 'document'}`}
        ariaLabel={`Open ${item.documentName || item.title || 'document'}`}
        onClick={() => onOpenItem(item)}
        styles={{ root: { border: 'none', background: 'transparent', padding: 0 } }}
      />*/
      <Icon
    iconName="InsertSignatureLine"
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
    }
  }
];

export default function DocumentsList(props: IDocumentsListProps): React.ReactElement {
  const [selectedStatus, setSelectedStatus] = React.useState<StatusFilter>('Pending');

  const filteredItems: IDocumentListItem[] = React.useMemo(() => {
    if (selectedStatus === 'Total') {
      return props.items;
    }

    const normalizedStatus: string = selectedStatus.toLowerCase();
    return props.items.filter((item: IDocumentListItem) =>
      (item.approvalStatus || '').toLowerCase().indexOf(normalizedStatus) !== -1
    );
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
      ) : (
        <DetailsList
          items={filteredItems}
          columns={getColumns(props.onOpenItem)}
          setKey="documentsList"
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          constrainMode={ConstrainMode.unconstrained}
          compact={true}
          ariaLabel="Signed documents list"
        />
      )}
    </div>
  );
}
