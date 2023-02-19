// eslint-disable-next-line @typescript-eslint/no-var-requires
const fetch = require('node-fetch');

type SharePointListColumn = {
  id: string;
  displayName: string;
  name: string;
  columnGroup: string;
  description?: string;
  indexed: boolean;
  readOnly: boolean;
  hidden: boolean;
  formula?: string;
  defaultValue?: {
    value: string | number | boolean;
  };
  validation?: {
    formula?: string;
    message?: string;
  };
  choice?: {
    allowTextEntry: boolean;
    choices: string[];
  };
  text?: {
    allowMultipleLines: boolean;
    appendChangesToExistingText: boolean;
    linesForEditing: number;
    maxLength: number;
  };
  personOrGroup?: {
    allowMultipleSelection: boolean;
  };
};

export const getListColumns = async (
  token: string,
  siteId: string,
  listId: string,
) => {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  const data = await response.json();
  const columns: { [key: string]: SharePointListColumn } = {};
  data.value.forEach((row: SharePointListColumn) => {
    columns[row.name] = row;
  });
  return columns;
};
