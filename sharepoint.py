"""Module to implement SharePoint functionality in bots

:Author: Sharath Kumar V K
:Contact: sharath.kumar.k.v@ericsson.com
:Date: 10th Oct, 2020

This module aims to provide clear logical interfaces for
implementing SharePoint functionalities in the bots.
Presently some of the developers might be editing the boiler
plate code provided in the `Automation Wiki Channel <https://teams.microsoft.com/l/entity/com.microsoft.teamspace.tab.wiki/tab::85b8d8f9-1bfa-452a-a118-a3c297d8fbde?context=%7B%22subEntityId%22%3A%22%7B%5C%22pageId%5C%22%3A7%2C%5C%22origin%5C%22%3A2%7D%22%2C%22channelId%22%3A%2219%3A33002c78ffa343ff841aa9e9d7f654e2%40thread.tacv2%22%7D&tenantId=92e84ceb-fbfd-47ab-be52-080c6b87953f>`_.
This module aims to mitigate that and provide a pythonic approach.

Also, if an individual bot needs to access multiple SharePoint sites,
multiple instances can be created for sites using the ``SharePointBuilderObject``,
which was not possible previously or it was cumbersome to edit the boiler-plate
for multiple sites.

Examples
--------
>>> import sharepoint
>>> sp = sharepoint.SharePointObjectBuilder()

If the configurations for a SharePoint site is already present in ``config.ini`` file.

>>> site_1 = sp('site_1_name')
>>> site_1.connection_status()
200
>>> site1.create_folder('Shared Documents/folder1', 'new_folder')
Folder creation attempt response: 200
Folder: new_folder created at Shared Documents/folder1

If you want to register new configurations of internal SharePoint site

>>> site2 = sp.register_site(site='new_site', client_id='1234..', client_secret='fylth..')
>>> site2.connection_status()
200

If you want to register new configurations of external SharePoint site

>>> site3 = sp.register_site(site='ext_site', client_id='456..', client_secret='xyz..',\
domain='telefonicacorp.sharepoint.com', tenant_id='9730...')
>>> site3.connection_status()
200

"""

from collections import namedtuple
from configparser import ConfigParser
from pathlib import Path
from typing import NamedTuple
from urllib import parse

import requests


# ConfigParser is used to read/write to 'config.ini' file
# parse is used to encode the url
# NamedTuple is used to make the type hinting clear wherever it is used

class SharePoint:
    """A class used to represent SharePoint object.

    ...

    Parameters
    ----------
    site : str
        Name of SharePoint site
    response: NamedTuple
        e.g. ``Response(status_code=200, token='eyb8fh...', domain='ericsson.sharepoint.com')``

    Attributes
    ----------
    site : str
        Name of SharePoint site
    access_token : str
        Token used for interactions with SharePoint site
    status_code : int
        Indicates the connection status with SharePoint site
    domain: str
        Domain information depending on internal or external site

    Warning
    -------
    This class should not be instantiated on its own.

    """

    def __init__(self, site: str, response: NamedTuple):
        self.site = site
        self.access_token = response.token
        self.status_code = response.status_code
        self.domain = response.domain
        self._set_headers()

    def __repr__(self):
        return F"{self.__class__.__name__}({self.site!r}, {self.domain!r})"

    def connection_status(self) -> int:
        """This method returns the connection status with SharePoint site

        Returns
        -------
        status_code : int
            HTTP status codes are returned

        """
        return self.status_code

    def _set_headers(self):
        self.headers = {"Accept": "application/json; odata=verbose",
                        "Content-Type": "application/json; odata=verbose",
                        "Authorization": F"Bearer {self.access_token}"}
        self.json_data = None

    def _get_metadata(self, folder_path: str) -> dict:
        url = F"https://{self.domain}/sites/{self.site}/_api/web/GetFolderByServerRelativeUrl('{folder_path}')/files"
        data = requests.get(url, headers=self.headers)
        self.json_data = data.json()

    def download_file(self, folder_path: str, file_name: str, path_to_save: str) -> Path:
        """Downloads file from specified SharePoint folder_path

        Parameters
        ----------
        folder_path : str or Path object
            Path where file is located. e.g. ``"Shared Documents/folder1/folder2"``
        file_name : str
            Name of the file to be downloaded with extension
        path_to_save : str
            Path where the file needs to be saved, can be a string or Path object

        Returns
        -------
        Path
            Path where file is downloaded

        Raises
        ------
        FileNotFoundError
            If file is not found at specified SharePoint path

        Warning
        -------
        Make sure your credentials have adequate permissions before calling the method

        """
        if not self.json_data:
            self._get_metadata(folder_path)
        file_name_data = self.json_data['d']['results']
        for key in file_name_data:
            if key['Name'] == file_name:
                url_to_file = F"https://{self.domain}/sites/{self.site}/_api/web/GetFolderByServerRelativeUrl('" \
                              F"{folder_path}')/Files('{file_name}')/$value"
                file_save_path = Path(path_to_save).joinpath(file_name)
                response = requests.get(url_to_file, headers=self.headers)
                with open(file_save_path, 'wb+') as file:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            file.write(chunk)
                            file.flush()
                print(F"\nStatus: {response.status_code}, File: {file_name} downloaded to {file_save_path.parent}")
                return file_save_path
        raise FileNotFoundError(F"{file_name} not found at: {folder_path}")

    def bulk_download(self, folder_path: str, path_to_save: str, files: list) -> dict:
        """Downloads multiple files at once

        Parameters
        ----------
        folder_path : str
            Path where files are located. e.g. ``"Shared Documents/folder1/folder2"``
        path_to_save : str
            Local path where files need to be saved
        files : list
            List of file names to be downloaded

        Returns
        -------
        dict
            A dictionary with keys as filenames and paths as values
        """
        if isinstance(files, list):
            file_paths = [self.download_file(folder_path, _, path_to_save) for _ in files]
            path_dict = {}
            for path in file_paths:
                path_dict[path.stem] = str(path)
        return path_dict

    def upload_file(self, folder_path, absolute_filepath):
        """Uploads file to specified SharePoint path.

        Parameters
        ----------
        folder_path : str or Path object
            path to which file has to be uploaded. e.g. ``"Shared Documents/folder1/folder2"``
        absolute_filepath : str or Path object
            absolute file path of the file to be uploaded. e.g. ``"C:\folder1\folder2\test.txt"``

        Warning
        -------
        Make sure your credentials have adequate permissions before calling the method

        Raises
        ------
        FileNotFoundError
            If file is not found at specified local directory

        """
        file_path = Path(absolute_filepath)
        if file_path.is_file():
            with file_path.open(mode='rb') as file:
                file_buffer = file.read()

            url = F"https://{self.domain}/sites/{self.site}/_api/web/getfolderbyserverrelativeurl('{folder_path}')/Files" \
                  F"/add(url='{file_path.name}', overwrite=true)"

            data = requests.post(url, headers=self.headers, data=file_buffer)
            if data.status_code == 200:
                print(F"\nStatus: {data.status_code}, File: {file_path.name} uploaded to {folder_path}")
            else:
                print(F"\nUpload attempt response: {data.status_code}")
                print(F"\n Upload response content: {data.content}")
        else:
            raise FileNotFoundError(F"{file_path.name} not found at: {file_path.parent}")

    def bulk_upload(self, folder_path: str, files_to_upload: list):
        """Uploads multiple files at once to a specified SharePoint path

        Parameters
        ----------
        folder_path: str
            Path where files need to uploaded. e.g. ``"Shared Documents/folder1/folder2"``
        files_to_upload: list
            List of file names to upload, file names. e.g. ``["C:\folder1\foo.txt", "C:\folder2\bar.xlsx"]``

        """
        for file in files_to_upload:
            self.upload_file(folder_path, file)

    def create_folder(self, folder_path: str, folder_name: str):
        """Creates folder at specified SharePoint path.

        Parameters
        ----------
        folder_path : str or Path object
            path at which folder needs to be created. e.g. ``"Shared Documents/folder1/folder2"``
        folder_name : str
            name of the folder which needs to be created

        Warnings
        --------
        Make sure your credentials have adequate permissions before calling the method

        """
        url = F'https://{self.domain}/sites/{self.site}/_api/web/folders'
        json = {"__metadata": {"type": "SP.Folder"},
                "ServerRelativeUrl": F"{folder_path}/{folder_name}"}
        response = requests.post(url, headers=self.headers, json=json)
        if response.status_code == 201:
            print(F"\nFolder creation attempt response: {response.status_code}")
            print(F"\nFolder: {folder_name} created at {folder_path}")
        else:
            print(F"\nFolder creation attempt response: {response.status_code}")
            print(F"\nFolder creation response content: {response.content}")


class SharePointObjectBuilder:
    """A class to build or instantiate SharePoint objects.

    ...

    Uses a configuration file ``config.ini`` to store information related to various SharePoint sites.


    Returns
    -------
    SharePoint object
        An instance of SharePoint object

    Raises
    ------
    FileNotFoundError
        if "config.ini" file is not found
    KeyError
        if config.ini doesn't contain a section related to requested SharePoint site
    ConnectionError
        if failed to generate access_token for a particular site

    Important
    ---------
    ``config.ini`` would look like this. With saved configurations, new SharePoint objects
    can be instantiated as shown in Examples. New configurations can be written using
    ``register_site`` method.
    |config file|

    """

    def __init__(self):
        self._instance = None

    def __call__(self, site: str):
        if site:
            response = self._check_site_in_config(site)
            if response.status_code == 200:
                self._instance = SharePoint(site, response)
                return self._instance
            else:
                raise ConnectionRefusedError(f"{response.status_code}: Connection failed, recheck configuration")

    def __repr__(self):
        return F"{self.__class__.__name__}('{self._instance}')"

    def _check_site_in_config(self, site: str):
        self._read_config_file()
        if site in self._parser:
            response = self._authorize(site)
            return response
        else:
            raise KeyError(f"Configuration not found for {site}")

    def _read_config_file(self):
        self._parser = ConfigParser()
        if self._parser.read('config.ini'):
            pass
        else:
            raise FileNotFoundError('Configuration file not found')

    def _get_configs(self, site: str):
        _configuration = namedtuple('Configs', ['TENANT_ID', 'DOMAIN', 'CLIENT_ID', 'CLIENT_SECRET'])
        self._configs = _configuration(self._parser.get(site, 'TENANT_ID'),
                                       self._parser.get(site, 'DOMAIN'),
                                       self._parser.get(site, 'CLIENT_ID'),
                                       self._parser.get(site, 'CLIENT_SECRET'))

    def _authorize(self, site: str) -> NamedTuple:
        self._get_configs(site)
        url = F"https://accounts.accesscontrol.windows.net/{self._configs.TENANT_ID}/tokens/OAuth/2"
        encoded = parse.quote(self._configs.CLIENT_SECRET)
        payload = F"grant_type=client_credentials&client_id={self._configs.CLIENT_ID}%40{self._configs.TENANT_ID}&" \
                  F"client_secret={encoded}&resource=00000003-0000-0ff1-ce00-000000000000%2F" \
                  F"{self._configs.DOMAIN}%40{self._configs.TENANT_ID}"
        headers = {
            'content-type': "application/x-www-form-urlencoded",
            'cache-control': "no-cache",
            'postman-token': "db08fff1-63bf-1f8d-84c2-4f466cc49afc"
        }
        res = requests.request("POST", url, data=payload, headers=headers)
        response = namedtuple('Response', ['status_code', 'token', 'domain'])
        return response(res.status_code, res.json().get('access_token'), self._configs.DOMAIN)

    def register_site(self, *, site: str, client_id: str, client_secret: str, domain: str = None,
                      tenant_id: str = None):
        """Method to register the configurations of a new SharePoint site

        Parameters
        ----------
        site : str
            Name of the SharePoint site
        client_id : str
            Generated client id during app registration on SharePoint site
        client_secret : str
            Generated client secret during app registration on SharePoint site
        domain : :obj:`str`, optional
            None by default, should be specified only when the site is external (non-ericsson)
        tenant_id : :obj:`str`, optional
            None by default, should be specified only when the site is external (non-ericsson)

        Returns
        -------
        SharePoint object
            An instance of SharePoint object

        Note
        ----
        Accepts only keyword arguments

        """
        self._read_config_file()
        self._parser[site] = {'client_id': client_id,
                              'client_secret': client_secret,
                              }
        if domain and tenant_id:
            self._parser[site].update({'domain': domain, 'tenant_id': tenant_id})
        with open('config.ini', 'w') as file:
            self._parser.write(file)
        sharepoint_object = self.__call__(site)
        return sharepoint_object

