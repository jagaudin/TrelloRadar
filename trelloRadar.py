# -*- coding: utf-8 -*-


#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.

#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.

#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>.

#   Created on Thu Jan  4 17:12:14 2018

#   @author: Jacques Gaudin <jagaudin@gmail.com>


import requests
import json
import configparser
import webbrowser

from pathlib import Path
from bs4 import BeautifulSoup
from datetime import datetime
from itertools import groupby

import tkinter as tk
from tkinter import ttk

import clr
clr.AddReference('System.Threading')
clr.AddReference('System.Windows')
clr.AddReference('System.Windows.Forms')
import System.Windows.Forms as WinForms
from System.Threading import Thread, ThreadStart, ApartmentState
from System.Drawing import Size


class AuthDialog:
    """
    A class to present an authorization dialog to the user.
    If `API_key` is not given an API key and token are looked for, otherwise
    only a token.

    :param API_key: an Trello API key, default `None`
    :returns: an `AuthDialog` object
    """

    class FormBrowser(WinForms.Form):
        """
        A class to implement a basic browser based on `Windows.Forms`.
        The browser is specifically tuned to retrieve the user's Trello
        API key and get a token.

        :param API_key: an Trello API key, default `None`
        :returns: a `FormBrowser` object
        """

        token_url = ('https://trello.com/1/authorize?key={0}'
                     '&name={1}&expiration={2}&response_type=token&scope={3}')

        name, expiry, scope = 'TrelloRadar', 'never', 'read,write'

        token_success_string = ('You have granted  access to '
                                'your Trello information.')

        login_url = 'https://trello.com/login'
        login_redirect_url = 'https://trello.com/'
        API_key_url = 'https://trello.com/app-key'

        api_key_success_string = 'Developer API Keys'

        def __init__(self, API_key):

            self.API_key = API_key
            self.token = ''

            if API_key:
                self.target_url = self.token_url.format(
                        API_key, self.name, self.expiry, self.scope)
            else:
                self.target_url = self.login_url

            self.Text = 'Authorization'
            self.ClientSize = Size(800, 800)

            self.FormBorderStyle = WinForms.FormBorderStyle.FixedSingle
            self.MaximizeBox = False

            self.web_browser = WinForms.WebBrowser()
            self.web_browser.Dock = WinForms.DockStyle.Fill
            self.web_browser.ScriptErrorsSuppressed = True
            self.web_browser.IsWebBrowserContextMenuEnabled = False
            self.web_browser.WebBrowserShortcutsEnabled = False

            self.web_browser.DocumentCompleted += self.on_document_completed
            self.web_browser.DocumentCompleted += self.check_token

            if not API_key:
                self.web_browser.Navigated += self.on_navigated
                self.web_browser.DocumentCompleted += self.check_API_key

            self.web_browser.Visible = True
            self.web_browser.Navigate(self.target_url)

            self.Controls.Add(self.web_browser)

        def on_navigated(self, sender, args):
            """
            Signal handler to redirect to the API key URL on
            successful login
            """

            self.web_browser.Visible = True
            # redirect main user page on successful login
            if str(self.web_browser.Url) == self.login_redirect_url:
                self.web_browser.Stop()
                self.web_browser.Navigate(self.API_key_url)

        def on_document_completed(self, sender, args):
            """
            Signal handler to parse the html content of the page
            """

            content = self.web_browser.DocumentText
            self.soup = BeautifulSoup(content, 'html.parser')

        def check_API_key(self, sender, args):
            """
            Signal handler to retrieve API key from html content
            """

            try:
                if self.api_key_success_string in self.soup.find('h1').string:
                    self.web_browser.Visible = False
                    self.API_key = self.soup.find('input', id='key')['value']
                    self.target_url = self.token_url.format(
                            self.API_key, self.name, self.expiry, self.scope)
                    self.web_browser.Navigate(self.target_url)
            except:
                pass

        def check_token(self, sender, args):
            """
            Signal handler to retrieve token from html content
            """

            try:
                if self.token_success_string in self.soup.p.string:
                    self.token = self.soup.find('pre').string.strip()
                    self.Close()
            except:
                pass

    def __init__(self, API_key=''):

        def start():
            self.browser = AuthDialog.FormBrowser(API_key)
            WinForms.Application.Run(self.browser)

        # Create a new thread to run the process
        thread = Thread(ThreadStart(start))
        thread.SetApartmentState(ApartmentState.STA)
        thread.Start()
        thread.Join()

    @property
    def API_key(self):
        return self.browser.API_key

    @property
    def token(self):
        return self.browser.token


class TrelloRadarApp():
    """
    The main application class

    :returns: a `TrelloRadarApp` object
    """

    config_path = (Path.home() / 'AppData' / 'Local' / 'TrelloRadar' /
                   'settings.ini')

    search_strings = ['@me']

    time_f = '%Y-%m-%dT%H:%M:%S.%fZ'

    def __init__(self):
        self.boards_by_id = {}
        self.boards_by_name = {}
        self.today = datetime.today().date()

        self.get_config()
        self.setup_gui()
        self.send_querystring()
        self.root.mainloop()

    def get_config(self):
        """
        Reads the config file

        :returns: `None`
        """

        if not self.config_path.exists():
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            self.config_path.touch()

        self.config = configparser.ConfigParser()
        self.config.read(self.config_path)

        # Read 'auth' section
        if (self.config.has_section('auth') and
                self.config.has_option('auth', 'API key')):

            self.API_key = self.config['auth']['API key']
            if self.config.has_option('auth', 'token'):
                self.token = self.config['auth']['token']
                self.validate_credentials()
            else:
                self.get_token()
        else:
            self.get_API_key()

        # Read 'search' section
        if self.config.has_option('search', 'search strings'):
            search_strings = self.config['search']['search strings']
            self.search_strings = search_strings.split(';')
        else:
            self.config['search'] = {}
            self.config['search']['search strings'] = self.search_strings[0]

    def get_API_key(self):
        """
        Launches an `AuthDialog` instance to retrieve the API key and token

        :returns: `None`
        """

        auth_dialog = AuthDialog()
        self.API_key = auth_dialog.API_key
        self.token = auth_dialog.token

        self.config['auth'] = {}
        self.config['auth']['API key'] = self.API_key
        self.config['auth']['token'] = self.token
        self.save_config()

    def get_token(self):
        """
        Launches an `AuthDialog` instance to retrieve a token

        :returns: `None`
        """

        auth_dialog = AuthDialog(self.API_key)
        self.token = auth_dialog.token

        self.config['auth']['token'] = self.token
        self.save_config()

    def validate_credentials(self):
        """
        Attempts a connection to Trello with the credentials present.
        If connection is refused, gets new credentials.

        :returns: `None`
        """

        url = 'https://api.trello.com/1/members/me/'
        token_query = {
            'key': self.API_key,
            'token': self.token,
        }
        response = requests.get(url, params=token_query)
        if response.status_code == 200:
            return
        elif len(self.API_key) == 32 and 'invalid token' in response.text:
            self.get_token()
        else:
            self.get_API_key()

    def show_data(self, query_string):
        """
        Shows the cards matching `query_string` in a tree view.

        :parameter query_string: a search query
        :returns: `None`
        """

        self.todo_tree.delete(*self.todo_tree.get_children())

        cards = self.search_cards(query_string)
        cards = sorted(cards, key=lambda c: c['board']['name'])

        for c in cards:
            board_url = c['board']['url']
            if not self.todo_tree.exists(board_url):
                board_name = c['board']['name']
                self.todo_tree.insert('', 'end', board_url, text=board_name)
                self.todo_tree.item(board_url, open=True)
            if c['due']:
                due_date = datetime.strptime(c['due'], self.time_f).date()
                tags = ('overdue',) if self.today > due_date else ()
            else:
                due_date = ''
                tags = ()

            labels = ', '.join(label['name'] for label in c['labels'])

            self.todo_tree.insert(board_url, 'end', c['shortUrl'],
                                  text=c['name'],
                                  values=(due_date, labels), tags=tags)

    def search_cards(self, query_string, cards_limit='1000'):
        """
        Search Trello for cards matching `query_string`.

        :parameter query_string: a search query
        :parameter cards_limit: max. number of cards returned (default `1000`)
        :returns: a list of the cards found
        """

        search_url = 'https://api.trello.com/1/search'

        search_query = {
            'key': self.API_key,
            'token': self.token,
            'modelTypes': 'cards',
            'card_board': 'true',
            'board_fields': 'name,url',
            'query': query_string,
            'cards_limit': cards_limit,
        }
        response = requests.get(search_url, params=search_query)
        return response.json()['cards'] if response.status_code == 200 else []

    def send_querystring(self):
        """
        Records the search query and shows results

        :returns: `None`
        """

        query_string = self.entry.get()
        if not query_string:
            return

        if query_string not in self.entry['values']:
            self.entry['values'] = (query_string,) + self.entry['values']

        self.show_data(query_string)

    def clear_search(self):
        """
        Resets the previous search list to `['@me']`

        :returns: `None`
        """

        self.entry['values'] = ['@me']

    def link_tree(self, event):
        url = self.todo_tree.selection()[0]
        webbrowser.open(url)

    def on_entry_return(self, event):
        self.send_querystring()

    def on_closing(self):
        config_search_strings = ';'.join(s for s in self.entry['values'])
        self.config['search']['search strings'] = config_search_strings
        self.save_config()
        self.root.destroy()

    def save_config(self):
        self.config.write(self.config_path.open('w'))

    def setup_gui(self):
        """
        Prepares the GUI

        :returns: `None`
        """

        self.root = tk.Tk()
        self.root.protocol('WM_DELETE_WINDOW', self.on_closing)
        try:
            self.root.iconbitmap(default='icons/transparent.ico')
        except:
            print('Icon file not found')

        self.root.geometry('540x700')
        self.root.title('Trello Radar')

        self.mainframe = ttk.Frame(self.root, padding='3 3 3 3')
        self.mainframe.grid(column=0, row=0, columnspan=3, sticky='nwes')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.todo_tree = ttk.Treeview(self.mainframe,
                                      columns=('Due Date', 'Label'))
        self.todo_tree.pack(expand=True, fill='both')
        self.todo_tree.heading('#0', text='Task')
        self.todo_tree.column('#0', minwidth=200, width=390, stretch=True)
        self.todo_tree.heading('Due Date', text='Due Date')
        self.todo_tree.column('Due Date', minwidth=50, width=70, stretch=False)
        self.todo_tree.heading('Label', text='Label')
        self.todo_tree.column('Label', minwidth=50, width=60, stretch=False)

        self.todo_tree.tag_configure('overdue', foreground='red')

        self.todo_tree.bind('<Double-1>', self.link_tree)

        self.entry = ttk.Combobox(self.root, values=self.search_strings)
        self.entry.insert(0, self.search_strings[0])
        self.entry.grid(column=0, row=1, sticky='ew')
        self.entry.bind('<Return>', self.on_entry_return)

        self.clear_button = ttk.Button(self.root, text='Clear search',
                                       command=self.clear_search)
        self.clear_button.grid(column=1, row=1)

        self.refresh_button = ttk.Button(self.root, text='Refresh',
                                         command=self.send_querystring)
        self.refresh_button.grid(column=2, row=1)

if __name__ == '__main__':
    trello_todo_app = TrelloRadarApp()
