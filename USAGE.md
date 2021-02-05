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

## Delegate check

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