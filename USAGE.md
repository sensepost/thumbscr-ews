# Usage 

```
thumbscr-ews --help
Usage: thumbscr-ews [OPTIONS] COMMAND [ARGS]...

  thumsc-ews for Exchange Web Services
      by @_cablethief from @sensepost of @orangecyberdef

Options:
  -C, --config PATH          Path to an optional configuration file.
  -u, --username TEXT        The username to use.
  -p, --password TEXT        The password to use.
  -a, --user-agent TEXT      The User-Agent to use (Otherwise uses "thumbscr-
                             ews/0.0.1").

  -o, --outlook-agent        Set the User-Agent to an Outlook one (this is
                             still a static value that can be fingerprinted).

  --dump-config              Dump the effective configuration used.
  --exch-host TEXT           If you dont want to try autodicover set the
                             exchange host.

  -v, --verbose              Enables debugging information.
  -w, --table-width INTEGER  The maximum width used for table output
  --help                     Show this message and exit.

Commands:
  autodiscover   Authenticate and go through autodiscover.
  brute          Do a brute force.
  delegatecheck  Check if the current user has access to the provided...
  folders        Print exchange file structure.
  gal            Dump GAL using EWS.
  mail           Do things with emails.
  objects        Discover objects.
  version        Prints the current thumbsc-ews version
  yaml           Generate an example YAML configuration file

```

## Config

The config file is just a way to help shorten or keep credentials out of your command line. You can generate a sample one using: 

```
thumbscr-ews yaml -d test.yml
# An example YAML configuration for thumbscr-ews.

password: passw0rd
user_agent: Microsoft Office/14.0 (Windows NT 6.1; Microsoft Outlook 14.0.7145; Pro)
username: user@domain.com

Sample configuration file written to: test.yml
```

All the values in the yaml can be used in the command line as well. 

## AutoDiscover

Just a command to go through all the autodiscover steps to help debug if something is going wrong or to show you what endpoint thumbscr-ews is trying to use. 

```
thumbscr-ews -C config.yml autodiscover
```

Autodiscover may not always work and you will have to work it out for yourself, a good start is `outlook.office365.com`. You can specify that for all future commands using `--exch-host`.

## GAL

One of the most useful functions and one that I would often use, would be to dump the gal. This is pretty straight forward but does have some caveats. To dump the gal the following can be used:

```
thumbscr-ews -C config.yml --exch-host outlook.office365.com gal
```

the caveat is that it is printing results as it gets them, due to the way that the search works (Looking from aa to zz) this could mean many doubles so you should always `sort -u` at the end when you have your emails. 
you can also search for specific strings in the GAL with `--search`. 

## Delegatecheck

With a list of emails, you can provide them to thumbscr-ews to see if you may have further access. 

```
thumbscr-ews ... delegatecheck -l delegateemails.txt                          
[-] michael@sensepost.com - Failure inbox not accessible
[-] dominic@sensepost.com - Failure inbox not accessible
[-] leon@sensepost.com - Failure inbox not accessible
[+] Success delegate@sensepost.com - Could access inbox - Permissions: ['Author', 'Custom']
```

By default the folder you check your access against is the users inbox. You can use `--folder` to specify a folder.

```
thumbscr-ews ... delegatecheck -l delegateemails.txt -f 'PeoplePublicData'   [+] Success michael@sensepost.com - Could access Folder (PeoplePublicData) - Permissions: ['Reviewer']
[+] Success dominic@sensepost.com - Could access Folder (PeoplePublicData) - Permissions: ['Reviewer']
[+] Success leon@sensepost.com - Could access Folder (PeoplePublicData) - Permissions: ['Reviewer']
[-] Failure delegate@sensepost.com - No folder found
```

The `--full-tree` command just prints out where you have some sort of access, even its its just viewing that the folder exists. 

```
thumbscr-ews ... delegatecheck -l delegateemails.txt --full-tree
[+] Success michael@sensepost.com - Access to some folders.
root
└── PeoplePublicData
[+] Success dominic@sensepost.com - Access to some folders.
root
└── PeoplePublicData
[+] Success leon@sensepost.com - Access to some folders.
root
└── PeoplePublicData
[+] Success delegate@sensepost.com - Access to some folders.
root
├── AllContacts
├── AllContactsExtended
├── AllItems
├── ApplicationDataRoot
│   ├── 33b68b23-a6c2-4684-99a0-fa3832792226
...
```

## Mail

Plans are to implement more things here, but for now there is the ability to read mails and get the associated attachments. 

```
Usage: thumbscr-ews mail [OPTIONS] COMMAND [ARGS]...

  Do things with emails.

Options:
  --help  Show this message and exit.

Commands:
  getattachments  Download all the attachments from a Mail
  read            Search for mail in folder.
```

### Mail Read

```
Usage: thumbscr-ews mail read [OPTIONS]

  Search for mail in folder. Default Inbox. For printing mail in a nice
  manner. If you give it a folder without mail objects you may be sad. Check
  the objects command for just printing out what the library gives us.

Options:
  --id TEXT            Get the email with the corresponding ID
  -s, --search TEXT    Provide a query string based on:
                       https://docs.microsoft.com/en-us/exchange/client-
                       developer/web-service-reference/querystring-
                       querystringtype.

  --html               Retrieve the HTML version of mails, default is text.
  -f, --folder TEXT    Specify the folder to read from. Default is Inbox. eg:
                       "Top of Information Store/Archive"

  -l, --limit INTEGER  Limit the results returned to the most recent <amount>.
                       Default 100

  -d, --delegate TEXT  Read a different persons mailbox you have access to
  --help               Show this message and exit.
```

If you want to read mail from a specific folder you can list the top `--limit` from that folder, or `--search` for specific mails. 
Once listed, the mails will also output their ID which can be used to query using the `--id` flag in future. 

You may also access another users mailbox in the same way by providing the `--delegate` flag. 

### Mail Getattachments

```
Usage: thumbscr-ews mail getattachments [OPTIONS]

  Download all the attachments from a Mail

Options:
  --id TEXT            Get the email attachments with the corrisponding ID.
                       Saved as md5(id)-attachmentname.

  -s, --search TEXT    Provide a query string based on:
                       https://docs.microsoft.com/en-us/exchange/client-
                       developer/web-service-reference/querystring-
                       querystringtype.

  --path PATH          Where to save your attachments. Default is current
                       directory

  -f, --folder TEXT    Specify the folder to read from. Default is Inbox. eg:
                       "Top of Information Store/Archive"

  -l, --limit INTEGER  Limit the results returned to the most recent <amount>.
                       Default 100

  -d, --delegate TEXT  Read a different persons mailbox you have access to
  --help               Show this message and exit.
```

Getattachments works in the exact same way as `read` however it will download any attachments associated with a mail it retrieves. So you can collect the top 10 mails Attachments with `-l 10` or you can specify with an `--id` which I would recommend to be a bit more targeted. The `--search` command acts the same and searches for words in mails, however it will then download the attachments in the found mails. 

## Folders

```
Usage: thumbscr-ews folders [OPTIONS]

  Print exchange file structure.

Options:
  -s, --search TEXT    Search pattern to glob on. eg "Top of Information
                       Store*"

  -d, --delegate TEXT  Read a different persons mailbox you have access to
  --help               Show this message and exit.
```

Folders will print out the folder structure of the mailbox. The `--search` flag can be used to filter the output down a particular branch, so that rather than printing out the entire folder structure we can filter to areas that we are interested in rather than the entire root structure.

```
thumbscr-ews ... folders --search 'Top of Information Store'

Top of Information Store
├── Archive
├── Calendar
│   ├── Birthdays
│   └── United States holidays
├── Contacts
│   ├── Companies
│   ├── GAL Contacts
│   ├── Organizational Contacts
│   ├── PeopleCentricConversation Buddies
│   ├── Recipient Cache
│   ├── {06967759-274D-40B2-A3EB-D7F9E73727D7}
│   └── {A9E2BC46-B3A0-4243-B315-60D991004455}
├── Conversation Action Settings
├── Conversation History
...
```

## Objects

```
Usage: thumbscr-ews objects [OPTIONS]

  Discover objects. Printing out the objects the library finds. Not nice and
  clean like the mail option. More for exploration for future features.
  Hopefully the object has a string version.

Options:
  -f, --folder TEXT    Specify the folder to read from. Default is Inbox. eg:
                       "Top of Information Store/Archive"

  -l, --limit INTEGER  Limit the results returned to the most recent <amount>.
                       Default 100

  -d, --delegate TEXT  Read a different persons mailbox you have access to
  --help               Show this message and exit.
```

Not all folders contain only mails, for these things I chose to generically print the object retrieved from `exchangelib`. This allows for the support of random objects without too much overhead. 
as an example a `contact` object can be printed out without fancy frills as with mail which is a more core item. (This is not the GAL however, for that we need to use the ResolveNames call)

```
thumbscr-ews ... objects -f 'AllContacts'     
Object: Contact(mime_content=b'BEGIN:VCARD\r\nPROFILE:VCARD\r\nVERSION:3.0\r\nMAILER:Microsoft Exchange\r\nPRODID:Microsoft Exchange\r\nFN:Michael Kruger\r\nN:;;;;\r\nEMAIL;TYPE=INTERNET:michael.kruger@orangecyberdefense.com\r\nORG:;\r\nCLASS:PUBLIC\r\nADR;TYPE=WORK:;;;;;;\r\nLABEL;TYPE=WORK:\r\nADR;TYPE=HOME:;;;;;;\r\nADR;TYPE=POSTAL:;;;;;;\r\nREV;VALUE=DATE-TIME:2020-09-16T19:17:43,440Z\r\nEND:VCARD\r\n', ...
```