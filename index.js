const sppull = require('sppull').sppull;
const dotenv = require('dotenv');

dotenv.config();

const context = {
  siteUrl: process.env.SHAREPOINT_SITE,
  creds: {
    username: process.env.SHAREPOINT_USERNAME,
    password: process.env.SHAREPOINT_PASSWORD
  }
};

const options = {
  spRootFolder: process.env.SHAREPOINT_ROOT,
  recursive: true,
  folderStructureOnly: true,
  ignoreEmptyFolders: false
};

sppull(context, options).then(() => {
  console.log('success');
}).catch(err => {
  console.log(err);
});

