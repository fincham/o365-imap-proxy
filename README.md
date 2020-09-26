# o365-imap-proxy

Proxy IMAP connections to Exchange Online and handle XOauth2 transparently. This makes it easier to use an IMAP client that doesn't support XOauth2 yet (such as https://sylpheed.sraoss.jp/en/), and allows these clients to connect more securely to Exchange.

At some point Microsoft is going to retire what they call "Basic auth", and so clients will need to implement XOauth2.

This repository is mainly an experiment, but I'll be adding to it as I use it myself. I expect to add SMTP support in the near future as well.

## Configuration

Set up a new virtualenv with `msal` and `xdg` installed as per the `requirements.txt`.

A sample `.ini` file is supplied. Copy this to `~/.config/o365-imap-proxy.ini` (or wherevere your XDG config home is) and modify it to suit. You'll need at least a "client ID" from Azure AD.

Run `o365_imap_proxy.py --register` to log the proxy in to your O365 account. This should only need to be done once, though I am not totally sure that the token refresh flows in the proxy actually work. This will need further investigation.

## Troubleshooting

This program is quite experimental so don't expect to rely on it. Passing `--debug` will enable additional debugging.
