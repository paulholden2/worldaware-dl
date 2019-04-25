const async = require('async');
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

let options = {
  spRootFolder: '',
  dlRootFolder: '',
  recursive: true,
  folderStructureOnly: false,
  ignoreEmptyFolders: false
};

// Peek will download just the first folder-level, to test root folder names
if (process.env.SHAREPOINT_PEEK) {
  options.recursive = false;
  options.folderStructureOnly = true;
}

const downloads = [
  {
    remote: 'Client Contracts  Current',
    local: '\\\\stria-prod1\\CID01570 - WorldAware\\JID01215 - CaaS\\SharePoint Files\\Client'
  },
  {
    remote: 'NonDisclosure Agreements',
    local: '\\\\stria-prod1\\CID01570 - WorldAware\\JID01215 - CaaS\\SharePoint Files\\NDA'
  },
  {
    remote: 'Vendor Contracts',
    local: '\\\\stria-prod1\\CID01570 - WorldAware\\JID01215 - CaaS\\SharePoint Files\\Vendor'
  },
  {
    remote: 'Independent Contractor Agreements',
    local: '\\\\stria-prod1\\CID01570 - WorldAware\\JID01215 - CaaS\\SharePoint Files\\Contractor'
  },
  {
    remote: 'Partnership Agreements',
    local: '\\\\stria-prod1\\CID01570 - WorldAware\\JID01215 - CaaS\\SharePoint Files\\Partner'
  }
];

async.eachSeries(downloads, (dl, callback) => {
  options.spRootFolder = dl.remote;
  options.dlRootFolder = dl.local;
  console.log(`${dl.remote} => ${dl.local}`);
  sppull(context, options).then(() => {
    console.log('Done');
    callback();
  }).catch(callback);
}, err => {
  if (err) {
    console.log(err);
  }

  console.log('Completed all');
});
