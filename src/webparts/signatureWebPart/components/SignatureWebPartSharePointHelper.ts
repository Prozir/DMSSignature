import type { IApprovalUpdate, IRejectionUpdate } from './SignatureWebPartInterfaces';

export function getListApiPath(listName: string): { apiPath: string; displayName: string } {
  const trimmed = listName.trim();
  const normalizeServerRelativeUrl = (url: string): string => url.replace(/ /g, '%20');

  if (/^https?:\/\//i.test(trimmed)) {
    try {
      const listUrl = new URL(trimmed);
      const serverRelativeUrl = listUrl.pathname;
      return {
        apiPath: `web/GetList('${normalizeServerRelativeUrl(serverRelativeUrl).replace(/'/g, "''")}')`,
        displayName: serverRelativeUrl.split('/').pop() || trimmed
      };
    } catch {
      // fallback to title-based access
    }
  }

  if (trimmed.indexOf('/') === 0) {
    const serverRelativeUrl = normalizeServerRelativeUrl(trimmed);
    return {
      apiPath: `web/GetList('${serverRelativeUrl.replace(/'/g, "''")}')`,
      displayName: serverRelativeUrl.split('/').pop() || trimmed
    };
  }

  return {
    apiPath: `web/lists/getbytitle('${trimmed.replace(/'/g, "''")}')`,
    displayName: trimmed
  };
}

export function getApprovalUpdate(currentStatus?: string): IApprovalUpdate | undefined {
  const normalizedStatus = (currentStatus || '').trim();

  if (normalizedStatus === 'L1 Pending For Signature') {
    return {
      approvalStatus: 'L1 Signed',
      comments: 'Approved by Level 1'
    };
  }

  if (normalizedStatus === 'L2 Pending For Signature') {
    return {
      approvalStatus: 'L2 Signed',
      comments: 'Approved by L2'
    };
  }

  return undefined;
}

export function getRejectionUpdate(currentStatus?: string): IRejectionUpdate | undefined {
  const normalizedStatus = (currentStatus || '').trim();

  if (normalizedStatus === 'L1 Pending For Signature') {
    return {
      approvalStatus: 'L1 Rejected'
    };
  }

  if (normalizedStatus === 'L2 Pending For Signature') {
    return {
      approvalStatus: 'L2 Rejected'
    };
  }

  return undefined;
}

export function getFileNameFromServerRelativeUrl(serverRelativeUrl?: string): string | undefined {
  if (!serverRelativeUrl) {
    return undefined;
  }

  return serverRelativeUrl.substring(serverRelativeUrl.lastIndexOf('/') + 1) || undefined;
}

export function normalizeServerRelativeUrlForApi(serverRelativeUrl: string): string {
  const normalizedSegments = serverRelativeUrl
    .trim()
    .split('/')
    .map((segment: string, index: number) => {
      if (index === 0 && segment === '') {
        return '';
      }

      let decodedSegment = segment;

      try {
        decodedSegment = decodeURIComponent(segment);
      } catch {
        // Keep the original segment if it is not valid URI encoding.
      }

      return encodeURIComponent(decodedSegment);
    })
    .join('/');

  return normalizedSegments.replace(/'/g, "''");
}