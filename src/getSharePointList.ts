// eslint-disable-next-line @typescript-eslint/no-var-requires
const fetch = require('node-fetch');

type SharePointListItem = {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  createdBy: {
    user: {
      id: string;
      displayName: string;
    };
  };
  lastModifiedBy: {
    user: {
      id: string;
      displayName: string;
    };
  };
  fields: {
    [key: string]: any;
  };
  etag: string;
  contentType: {
    id: string;
  };
  driveItem: {
    id: string;
  };
  sharepointIds: {
    listId: string;
    listItemId: string;
  };
};

export const getSharePointList = async (
  token: string,
  siteId: string,
  listId: string,
): Promise<SharePointListItem[]> => {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  const data = await response.json();
  return data.value;
};
