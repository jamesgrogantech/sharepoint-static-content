// eslint-disable-next-line @typescript-eslint/no-var-requires
const fetch = require('node-fetch');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { mkdirSync, writeFileSync } = require('fs');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const sharp = require('sharp');

export const getImageData = async (
  token: string,
  siteId: string,
  imgId: string,
  folderPath: string,
) => {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/Site Assets/items/${imgId}/driveItem`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  mkdirSync(folderPath, { recursive: true });
  const data = await response.json();
  const downloadUrl = data['@microsoft.graph.downloadUrl'];
  const fileName = data.name;
  const response2 = await fetch(downloadUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  const filePath = `${folderPath}/${fileName.replace(/\.[^/.]+$/, '')}.webp`;
  const buffer = await response2.buffer();
  const image = sharp(buffer);
  const metadata = await image.metadata();

  if (
    metadata.width &&
    metadata.height &&
    (metadata.width > 2000 || metadata.height > 2000)
  ) {
    image.resize({ width: 2000, height: 2000, fit: 'inside' });
  }

  const resizedBuffer = await image.webp({ quality: 60 }).toBuffer();
  writeFileSync(filePath, resizedBuffer);
  return filePath;
};
