import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, ConstrainMode, IconButton, IColumn, SelectionMode, Spinner, SpinnerSize } from '@fluentui/react';

export interface IDocumentListItem {
  id: number;
  title?: string;
  documentName?: string;
  documentNumber?: string;
  trader?: string;
  accountCode?: string;
  approvalStatus?: string;
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
    onRender: (item: IDocumentListItem) => (
      <IconButton
        iconProps={{ iconName: 'OpenInNew' }}
        title={`Open ${item.documentName || item.title || 'document'}`}
        ariaLabel={`Open ${item.documentName || item.title || 'document'}`}
        onClick={() => onOpenItem(item)}
      />
    )
  }
];

export default function DocumentsList(props: IDocumentsListProps): React.ReactElement {
  if (props.isLoading) {
    return <Spinner label="Loading signed documents..." size={SpinnerSize.medium} />;
  }

  if (props.items.length === 0) {
    return <div>No signed documents with PDF attachments were found.</div>;
  }

  return (
    <DetailsList
      items={props.items}
      columns={getColumns(props.onOpenItem)}
      setKey="documentsList"
      selectionMode={SelectionMode.none}
      layoutMode={DetailsListLayoutMode.justified}
      constrainMode={ConstrainMode.unconstrained}
      compact={true}
      ariaLabel="Signed documents list"
    />
  );
}
