import hashlib
import os
import time
import uuid
import json
import click
import yaml as yamllib
from slugify import slugify

from thumbscrews.tbestate import tbestate
from thumbscrews.__init__ import __version__

import exchangelib
import logging

from hashlib import md5

from exchangelib import Credentials, Account
from exchangelib import Account, DistributionList
from exchangelib import discover, BaseProtocol
from exchangelib.indexed_properties import EmailAddress
from exchangelib import Account, EWSDateTime, FolderCollection, Q, Message
from exchangelib import Account, FileAttachment, ItemAttachment, Message, CalendarItem, HTMLBody

# This handler will pretty-print and syntax highlight the request and response XML documents
from exchangelib.util import PrettyXmlHandler


@click.group()
@click.option('--config', '-C', type=click.Path(), help='Path to an optional configuration file.')
@click.option('--username', '-u', help='The username to use.')
@click.option('--password', '-p', help='The password to use.')
@click.option('--user-agent', '-a', help='The User-Agent to use (Otherwise uses "thumbscr-ews/0.0.1").')
@click.option('--outlook-agent', '-o', is_flag=True, help='Set the User-Agent to an Outlook one (this is still a static value that can be fingerprinted).')
# @click.option('--exchange', '-e', help='The exchange endpoint to use. eg: https://outlook.office365.com/EWS/Exchange.asmx')
@click.option('--dump-config', is_flag=True, help='Dump the effective configuration used.')
@click.option('--verbose', '-v', is_flag=True, help="Enables debugging information.")
@click.option('--table-width', '-w', type=click.INT, help='The maximum width used for table output', default=120)
def cli(config, username, password, dump_config, verbose, user_agent, outlook_agent, table_width):
    """
        \b
        thumsc-ews for Exchange Web Services
            by @_cablethief from @sensepost
    """
    # logging.basicConfig(level=logging.WARNING)

    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    BaseProtocol.USERAGENT = "thumbscr-ews/" + \
        __version__ + " (" + BaseProtocol.USERAGENT + ")"

    if outlook_agent and user_agent:
        print("CANNOT USE TWO USERAGENTS AT ONCE!!!")
        print("Please choose only --user-agent or --outlook-agent.")
        quit()

    if outlook_agent:
        user_agent = "Microsoft Office/14.0 (Windows NT 6.1; Microsoft Outlook 14.0.7145; Pro)"

    if user_agent:
        BaseProtocol.USERAGENT = user_agent

    # set the mq configuration based on the configuration file
    if config is not None:
        with open(config) as f:
            config_data = yamllib.load(f, Loader=yamllib.FullLoader)
            tbestate.dictionary_updater(config_data)

    # set configuration based on the flags this command got
    tbestate.dictionary_updater(locals())

    # If we should be dumping configuration, do that.
    if dump_config:
        click.secho('Effective configuration for this run:', dim=True)
        click.secho('-------------------------------------', dim=True)
        click.secho(f'Username:               {tbestate.username}', dim=True)
        click.secho(f'Password:               {tbestate.password}', dim=True)
        click.secho(f'User-Agent:             {tbestate.user_agent}', dim=True)
        click.secho('-------------------------------------\n', dim=True)


@cli.command()
def version():
    """
        Prints the current thumbsc-ews version
    """

    click.secho(f'thumbsc-ews version {__version__}')


@cli.command()
@click.option('--destination', '-d', default='config.yml', show_default=True,
              help='Destination filename to write the sample configuration file to.')
def yaml(destination):
    """
        Generate an example YAML configuration file
    """

    # Don't be a douche and override an existing configuration
    if os.path.exists(destination):
        click.secho(
            f'The configuration file \'{destination}\' already exists.', fg='yellow')
        if not click.confirm('Override?'):
            click.secho('Not writing a new sample configuration file')
            return

    config = {
        'username': 'user@domain.com',
        'password': 'passw0rd',
        'user_agent': 'Microsoft Office/14.0 (Windows NT 6.1; Microsoft Outlook 14.0.7145; Pro)'
    }

    click.secho('# An example YAML configuration for thumbscr-ews.\n', dim=True)
    click.secho(yamllib.dump(config, default_flow_style=False), bold=True)

    try:
        with open(destination, 'w') as f:
            f.write('# A thumbscr-ews configuration file\n')
            f.write(yamllib.dump(config, default_flow_style=False))

        click.secho(
            f'Sample configuration file written to: {destination}', fg='green')

    except Exception as ye:
        click.secho(
            f'Failed to write sample configuration file with error: {str(ye)}', fg='red')


@cli.command()
@click.option('--verbose', '-v', is_flag=True, help='This gives more information from autodiscover.')
def autodiscover(verbose):
    """
        Authenticate and go through autodiscover.
    """

    tbestate.validate(['username', 'password'])

    credentials = Credentials(
        tbestate.username, tbestate.password)

    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    primary_address, protocol = discover(
        tbestate.username, credentials)

    click.secho(
        f'Autodiscover results:', bold=True, fg='yellow')

    click.secho(f'{primary_address.user}', fg='bright_green')
    click.secho(f'{protocol}', fg='bright_green')


@cli.group()
def mail():
    """
        Do things with emails. Like print and send.
    """

    pass


@mail.command()
@click.option('--id', help='Get the email with the corrisponding ID')
@click.option('--search', '-s', help='Provide a query string based on: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype.')
@click.option('--html', is_flag=True, help='Retrieve the HTML version of mails, default is text.')
@click.option('--folder', '-f', help='Specify the folder to read from. Default is Inbox. eg: "Top of Information Store/Archive"')
@click.option('--limit', '-l', type=click.INT, help='Limit the results returned to the most recent <amount>. Default 100')
def print(search, html, limit, folder, id):
    """
        Search for mail in folder. Default Inbox.
        For printing mail in a nice manner. 
        If you give it a folder without mail objects you may be sad.
        Check the objects command for just printing out what the library gives us.
    """

    if limit:
        max = limit
    else:
        max = 100

    credentials = Credentials(
        tbestate.username, tbestate.password)
    account = Account(tbestate.username,
                      credentials=credentials, autodiscover=True)

    if folder:
        current_folder = account.root.glob(folder)
    else:
        current_folder = account.inbox

    if search:
        # mails = account.inbox.filter(Q(body__icontains=search) | Q(subject__icontains=search))
        # mails = account.inbox.filter(Q(body__icontains=search))
        mails = current_folder.filter(search).order_by(
            '-datetime_received')
    else:
        if id:
            mails = [current_folder.get(
                id=id)]
        else:
            mails = current_folder.all().order_by('-datetime_received')[:max]

    for item in mails:
        click.secho(f'Subject: {item.subject}', fg='bright_blue', bold=True)
        click.secho(f'Sender: {item.sender}', fg='bright_cyan')
        click.secho(f'ReceivedBy: {item.received_by}', fg='cyan')
        click.secho(f'ID: {item.id}\n', fg='bright_magenta')
        # click.secho(f'{item.datetime_received}', fg='green', bold=True)
        if html:
            click.secho(f'Body:\n\n{item.body}\n', fg='white')
        else:
            click.secho(f'Body:\n\n{item.text_body}\n', fg='white')
        if item.has_attachments:
            click.secho(f'Attachments:', fg='yellow', dim=True)
            for attach in item.attachments:
                click.secho(f'{attach.name} - {attach.content_type}',
                            fg='bright_yellow', dim=True)

        click.secho(f'-------------------------------------\n', dim=True)


@mail.command()
@click.option('--id', help='Get the email attachments with the corrisponding ID. Saved as md5(id)-attachmentname.')
@click.option('--search', '-s', help='Provide a query string based on: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype.')
@click.option('--path', type=click.Path(), help='Where to save your attachments. Default is current directory')
@click.option('--folder', '-f', help='Specify the folder to read from. Default is Inbox. eg: "Top of Information Store/Archive"')
@click.option('--limit', '-l', type=click.INT, help='Limit the results returned to the most recent <amount>. Default 100')
def getattachments(id, folder, path, search, limit):
    """
        Download all the attachments from a Mail
    """

    credentials = Credentials(
        tbestate.username, tbestate.password)
    account = Account(tbestate.username,
                      credentials=credentials, autodiscover=True)

    if limit:
        max = limit
    else:
        max = 100

    if folder:
        current_folder = account.root.glob(folder)
    else:
        current_folder = account.inbox

    if search:
        # mails = account.inbox.filter(Q(body__icontains=search) | Q(subject__icontains=search))
        # mails = account.inbox.filter(Q(body__icontains=search))
        mails = current_folder.filter(search).order_by(
            '-datetime_received')
    else:
        if id:
            mails = [current_folder.get(
                id=id)]
        else:
            mails = current_folder.all().order_by('-datetime_received')[:max]

    if not path:
        path = os.getcwd()

    for item in mails:
        uniqifiyer = md5(item.id.encode("utf-8")).hexdigest()
        click.secho(f'Subject: {item.subject}', fg='bright_blue', bold=True)
        click.secho(f'Sender: {item.sender}', fg='bright_cyan')
        click.secho(f'ReceivedBy: {item.received_by}', fg='cyan')
        click.secho(f'ID: {item.id}', fg='bright_magenta')
        click.secho(f'IDHash: {uniqifiyer}\n', fg='bright_magenta')
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                local_path = os.path.join(
                    path, uniqifiyer + '-' + attachment.name)
                with open(local_path, 'wb') as f, attachment.fp as fp:
                    buffer = fp.read(1024)
                    while buffer:
                        f.write(buffer)
                        buffer = fp.read(1024)
                click.secho(f'Saved attachment to {local_path}', fg='green')

        click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--search', '-s', help='Search pattern to glob on. eg "Top of Information Store*"')
# @click.option('--html', is_flag=True, help='Retrieve the HTML version of mails, default is text.')
# @click.option('--limit', '-l', type=click.INT, help='Limit the results returned to the most recent <amount>')
def folders(search):
    """
        Print exchange file structure.
    """
    credentials = Credentials(
        tbestate.username, tbestate.password)
    account = Account(tbestate.username,
                      credentials=credentials, autodiscover=True)

    account.root.refresh()

    if search:
        for branches in account.root.glob(search):
            click.secho(f'{branches.tree()}')
            click.secho(f'-------------------------------------\n', dim=True)
    else:
        click.secho(f'{account.root.tree()}')
        click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--folder', '-f', help='Specify the folder to read from. Default is Inbox. eg: "Top of Information Store/Archive"')
@click.option('--limit', '-l', type=click.INT, help='Limit the results returned to the most recent <amount>. Default 100')
def objects(limit, folder):
    """
        Discover objects.
        Printing out the objects the library finds. 
        Not nice and clean like the mail option. 
        More for exploration for future features. 
        Hopefully the object has a string version. 
    """

    if limit:
        max = limit
    else:
        max = 100

    credentials = Credentials(
        tbestate.username, tbestate.password)
    account = Account(tbestate.username,
                      credentials=credentials, autodiscover=True)

    if folder:
        current_folder = account.root.glob(folder)
    else:
        current_folder = account.inbox

    mails = current_folder.all()[:max]

    for item in mails:
        click.secho(f'Object: {item}', fg='white')

        click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--folder', '-f', help='Specify the folder to read from in the contacts dir. Default is GAL Contacts. ')
@click.option('--limit', '-l', type=click.INT, help='Limit the results returned to the most recent <amount>. Default All')
@click.option('--amount', '-a',  is_flag=True, help='Print the amount of contacts in the folder.')
def contacts(limit, folder, amount):
    """
        Discover Contacts. 
    """

    if limit:
        max = limit
    else:
        max = 100

    credentials = Credentials(
        tbestate.username, tbestate.password)
    account = Account(tbestate.username,
                      credentials=credentials, autodiscover=True)

    if folder:
        current_folder = account.contacts.glob(folder)
    else:
        current_folder = account.contacts / 'GAL Contacts'

    total = current_folder.all().count()

    if amount:
        click.secho(f'Amount: {total}', fg='bright_blue', bold=True)

    if limit:
        max = limit
    else:
        max = total

    mails = current_folder.all().order_by('-datetime_received')[:max]

    for item in mails:
        click.secho(f'Object: {item}', fg='white')

        click.secho(f'-------------------------------------\n', dim=True)


if __name__ == '__main__':
    cli()
