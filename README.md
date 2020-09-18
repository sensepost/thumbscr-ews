
<h1 align="center">
  <br>
    📧 Thumbscr-EWS
  <br>
  <br>
</h1>

<h4 align="center">A wrapper around the amazing <a href="https://ecederstrand.github.io/exchangelib/">exchangelib</a> to do some common EWS operations.</h4>
<p align="center">
  <a href="https://twitter.com/_cablethief"><img src="https://img.shields.io/badge/twitter-%40_cablethief-blue.svg" alt="@_cablethief" height="18"></a>
</p>
<br>

## Introduction

`thumbscr-ews` is a small Python utility used with Exchange Web Services. Using `thumbscr-ews`, it is possible to read and search through mail, retrieve the Global Address List, and download attachments. 

## Features

With `thumbscr-ews`, you can:

- Read emails. 
- Search for strings in emails. (kinda)
- Download attachments for emails. 
- Dump the GAL
- Given that it is using EWS if Legacy authentication is enabled 2FA [can be bypassed](https://www.blackhillsinfosec.com/bypassing-two-factor-authentication-on-owa-portals/). 


## Installation - docker

A docker container for `thumbscr-ews` exists and can be used with:

```text
docker run --rm -ti cablethief/thumbscr-ews
```

## Installation - host

Finally, `thumbscr-ews` itself can be installed with:

```bash

```
