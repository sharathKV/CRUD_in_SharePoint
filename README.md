#CRUD Operations with SharePoint site using RESTful apis  

Upload/download files, create/delete folders at a SharePoint site through cli  
  

Example Usage:  

`import sharepoint`
`sp = sharepoint.SharePointObjectBuilder()`  

##If the configurations for a SharePoint site is already present in ``config.ini`` file.

`site_1 = sp('site_1_name')`
`site_1.connection_status()`
200
`site1.create_folder('Shared Documents/folder1', 'new_folder')`  
Folder creation attempt response: 200
Folder: new_folder created at Shared Documents/folder1

##If you want to register new configurations of internal SharePoint site  

`site2 = sp.register_site(site='new_site', client_id='1234..', client_secret='fylth..')`
`site2.connection_status()`  
200

##If you want to register new configurations of external SharePoint site  

`site3 = sp.register_site(site='ext_site', client_id='456..', client_secret='xyz..',\`
`domain='telefonicacorp.sharepoint.com', tenant_id='9730...')`
`site3.connection_status()`
200


