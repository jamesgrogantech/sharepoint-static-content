// eslint-disable-next-line @typescript-eslint/no-var-requires
const fetch = require('node-fetch');

export type SharePointUser = {
  name: string;
  email: string;
};

export const getUserListId = async (token: string, siteId: string) => {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=system,id,name`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  const data = await response.json();
  const userList = data.value.find((list: any) => list.name === 'users');
  return userList.id as string;
};

export const getUserList = async (token: string, siteId: string) => {
  const userListId = await getUserListId(token, siteId);
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${userListId}/items?expand=fields`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  const data = await response.json();
  const users: { [key: string]: SharePointUser } = {};
  data.value?.forEach((user: any) => {
    users[user.id] = {
      name: user.fields.Title,
      email: user.fields.EMail,
    };
  });
  return users;
};
