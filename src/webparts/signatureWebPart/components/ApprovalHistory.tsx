import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, Spinner, SpinnerSize } from '@fluentui/react';
import styles from './SignatureWebPart.module.scss';
import type { IDocumentListItem } from './DocumentsList';
import type { IApprovalHistoryItem } from './SignatureWebPartInterfaces';

export interface IApprovalHistoryProps {
  isOpen: boolean;
  isLoading: boolean;
  errorMessage: string;
  document?: IDocumentListItem;
  items: IApprovalHistoryItem[];
  onDismiss: () => void;
}

const getDateTimeValue = (value?: string): string => value ? new Date(value).toLocaleString() : '-';
const getDocumentName = (document?: IDocumentListItem): string => document ? document.documentName || document.title || 'document' : 'document';

export default function ApprovalHistory(props: IApprovalHistoryProps): React.ReactElement {
  const itemstoview = props.items.filter((item:IApprovalHistoryItem) => (item.approvalStatus || '').trim().length>0);
  //<span>Document Name: {props.document.documentNumber || '-'}</span>
  return (
    <Dialog
      hidden={!props.isOpen}
      onDismiss={props.onDismiss}
      dialogContentProps={{
        title: `Approval history : ${props.document?.documentNumber || ''}` //Approval history : ${getDocumentName(props.document)}`
      }}
      modalProps={{ isBlocking: false }}
      minWidth="min(90vw, 900px)"
      maxWidth="900px"
    >
      <div className={styles.approvalHistoryContent}>
        {props.document && (
          <div className={styles.approvalHistoryIdentity}>
            <span>Document Name : {getDocumentName(props.document)}</span>
          </div>
        )}

        {props.isLoading && <Spinner label="Loading approval history..." size={SpinnerSize.medium} />}

        {!props.isLoading && props.errorMessage && (
          <div className={styles.approvalHistoryMessage} role="alert">{props.errorMessage}</div>
        )}

        
        {!props.isLoading && !props.errorMessage && itemstoview.length === 0 && (
          <div className={styles.approvalHistoryMessage}>No approval history was found for this document.</div>
        )}

        
         {!props.isLoading && !props.errorMessage && itemstoview.length > 0 && ( 
          <div className={styles.approvalHistoryTableWrap}>
            <table className={styles.approvalHistoryTable}>
              <thead>
                <tr>
                  <th scope="col">Status</th>
                  <th scope="col">Approval created for</th>
                  <th scope="col">Action taken by</th>
                  <th scope="col">Action date</th>
                  <th scope="col">Comments</th>
                </tr>
              </thead>
              <tbody>
                {itemstoview.map((item: IApprovalHistoryItem) => (
                  <tr key={item.id}>
                    <td>{item.approvalStatus || '-'}</td>
                    <td>{item.approvalRecipient || '-'}</td>
                    <td>{item.actionTakenBy || '-'}</td>
                    <td>{getDateTimeValue(item.actionTakenOn)}</td>
                    <td>{item.comments || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
      <DialogFooter>
        <DefaultButton text="Close" onClick={props.onDismiss} />
      </DialogFooter>
    </Dialog>
  );
}
