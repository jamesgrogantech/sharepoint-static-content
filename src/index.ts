#!/usr/bin/env node
// eslint-disable-next-line @typescript-eslint/no-var-requires
require('dotenv').config();
import { getImageData } from './getImageData';
import { getListColumns } from './getListColumns';
import { getSharePointList } from './getSharePointList';
import { getAccessToken } from './getToken';
import { getUserList } from './getUserList';
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { mkdirSync, writeFileSync } = require('fs');
console.log(`🔑 Authenticating SharePoint Online`);

const generate = async () => {
  mkdirSync('./assets', { recursive: true });
  const token = await getAccessToken();

  const [rows, columns, userList] = await Promise.all([
    getSharePointList(
      token,
      process.env.SSC_SITE_ID as string,
      process.env.SSC_LIST_ID as string,
    ),
    getListColumns(
      token,
      process.env.SSC_SITE_ID as string,
      process.env.SSC_LIST_ID as string,
    ),
    getUserList(token, process.env.SSC_SITE_ID as string),
  ]);

  const columnTypes: {
    [key: string]: {
      type:
        | 'text'
        | 'number'
        | 'choice'
        | 'multipleChoice'
        | 'image'
        | 'person';
      options?: string[];
    };
  } = {};
  Object.entries(rows[0].fields).forEach(([key, value]) => {
    if (columns[key]) {
      if (columns[key].text) {
        columnTypes[key] = {
          type: 'text',
        };
      } else if (columns[key].choice) {
        if (typeof value === 'string') {
          columnTypes[key] = {
            type: 'choice',
            options: columns[key].choice?.choices || [],
          };
        } else {
          columnTypes[key] = {
            type: 'multipleChoice',
            options: columns[key].choice?.choices || [],
          };
        }
      } else if (typeof value === 'number') {
        columnTypes[key] = {
          type: 'number',
        };
      } else {
        try {
          JSON.parse(value);
          columnTypes[key] = {
            type: 'image',
          };
        } catch (e) {
          columnTypes[key] = {
            type: 'text',
          };
        }
      }
    }
    if (
      columns[key.replace('LookupId', '')] &&
      'personOrGroup' in columns[key.replace('LookupId', '')]
    ) {
      columnTypes[key] = {
        type: 'person',
      };
    }
  });
  const data: {
    [key: string]: {
      [key: string]: string | number | boolean | string[] | undefined;
    };
  }[] = [];

  const images: Promise<string>[] = [];
  for (const row of rows) {
    const rowData: {
      [key: string]: string | number | boolean | string[] | undefined;
    } = {};
    const fields = Object.entries(row.fields);
    for (const [key, value] of fields) {
      let newValue = value;
      if (columnTypes[key]) {
        if (columnTypes[key].type === 'image') {
          try {
            const folderPath = `${process.env.SSC_FOLDER_PATH}/images`;
            const json = JSON.parse(value);
            if (json?.id && json?.fileName) {
              newValue = `${folderPath}/${json.fileName}`;
              images.push(
                getImageData(
                  token,
                  process.env.SSC_SITE_ID as string,
                  json.id,
                  folderPath,
                ),
              );
            }
          } catch (e) {}
        } else if (columnTypes[key].type === 'person') {
          newValue = userList[value];
          rowData[key.replace('LookupId', '')] = newValue;
          continue;
        }
      }
      rowData[key] = newValue;
    }
    data.push(rowData as any);
  }
  console.log('🖼️  Optimising Images...');
  await Promise.all(images);
  console.log(`💾 Saving content to: ${process.env.SSC_FOLDER_PATH}/data.json`);
  writeFileSync(
    `${process.env.SSC_FOLDER_PATH}/data.json`,
    JSON.stringify(data),
  );
};

generate();
