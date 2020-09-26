#!/usr/bin/python3

"""
o365-imap-proxy
Proxy IMAP connections to Exchange Online and handle XOauth2 transparently.

Michael Fincham <michael@hotplate.co.nz> 2020-09-26
"""

import argparse
import atexit
import configparser
import imaplib
import json
import logging
import os
import select
import socketserver
import sys
import hmac

from xdg import XDG_CACHE_HOME, XDG_CONFIG_HOME
import msal

buffer_size = 8192


class ImapHandler(socketserver.StreamRequestHandler):
    def handle(self):
        logging.info(
            "New connection from %s:%s"
            % (self.client_address[0], self.client_address[1])
        )
        if "access_token" not in result:
            logging.error("No valid access token to log in to Exchange")
            return

        self.imap_conn = imaplib.IMAP4_SSL(
            config["o365"].get("hostname", "outlook.office365.com")
        )
        self.imap_conn.debug = 4 if args.debug else 0
        self.imap_conn.authenticate("XOAUTH2", lambda x: auth_string)
        if self.imap_conn.select():
            logging.info(
                "%s:%s Upstream Exchange authenticated"
                % (self.client_address[0], self.client_address[1])
            )
        else:
            logging.error(
                "%s:%s failed to authenticate to upstream Exchange"
                % (self.client_address[0], self.client_address[1])
            )
            return

        new_capabilities = [
            m for m in self.imap_conn.capabilities if not m.upper().startswith("AUTH=")
        ] + ["AUTH=LOGIN"]
        state = 0  # 0 = client pre-auth, 1 = client authenticated
        self.upstream_socket = self.imap_conn.socket()
        self.downstream_socket = self.request

        logging.debug(
            "%s:%s New capabilities: %s"
            % (
                self.client_address[0],
                self.client_address[1],
                " ".join(new_capabilities),
            )
        )
        self.wfile.write(
            (
                "* OK [CAPABILITY %s] Hotplate O365 IMAP Proxy\r\n"
                % " ".join(new_capabilities)
            ).encode("utf-8")
        )

        while True:
            reading, writing, dud = select.select(
                [self.downstream_socket, self.upstream_socket],
                [],
                [self.downstream_socket, self.upstream_socket],
                0.1,
            )
            in_bytes = True

            if self.downstream_socket in reading:
                if state == 1:  # post-auth
                    in_bytes = self.downstream_socket.recv(buffer_size)
                    self.upstream_socket.sendall(in_bytes)
                elif state == 0:  # pre-auth
                    in_line = [
                        m.decode("utf-8") for m in self.rfile.readline().strip().split()
                    ]
                    identifier = in_line[0]
                    command = in_line[1].upper()
                    if command == "CAPABILITY":
                        self.wfile.write(
                            ("* CAPABILITY %s\r\n" % " ".join(new_capabilities)).encode(
                                "utf-8"
                            )
                        )
                        self.wfile.write(
                            ("%s OK CAPABILITY completed.\r\n" % identifier).encode(
                                "utf-8"
                            )
                        )
                    elif command == "LOGIN":
                        if hmac.compare_digest(in_line[-1], config["imap"]["password"]):
                            self.wfile.write(
                                (
                                    "%s OK [CAPABILITY %s] Logged in.\r\n"
                                    % (identifier, " ".join(new_capabilities))
                                ).encode("utf-8")
                            )
                            state = 1
                            logging.info(
                                "%s:%s Downstream IMAP authenticated, proceeding to proxy"
                                % (self.client_address[0], self.client_address[1])
                            )
                    elif command == "LOGOUT":
                        self.wfile.write(
                            "* BYE IMAP Proxy is done with ya!\r\n".encode("utf-8")
                        )
                        self.wfile.write(
                            ("%s OK LOGOUT completed." % identifier).encode("utf-8")
                        )

            elif self.upstream_socket in reading and state == 1:
                in_bytes = self.upstream_socket.recv(buffer_size)
                self.downstream_socket.sendall(in_bytes)

            if not in_bytes or dud:
                if self.upstream_socket in dud:
                    logging.info("Upstream (O365) socket error, closing connection")
                if self.downstream_socket in dud:
                    logging.info("Downstream (IMAP) socket error, closing connection")
                self.upstream_socket.close()
                self.downstream_socket.close()
                return

    def finish(self):
        logging.info(
            "Connection from %s:%s completed"
            % (self.client_address[0], self.client_address[1])
        )


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--configuration",
        type=argparse.FileType("r"),
        default="%s/o365-imap-proxy.ini" % XDG_CONFIG_HOME,
    )
    parser.add_argument("--register", default=False, action="store_true")
    parser.add_argument("--debug", default=False, action="store_true")
    args = parser.parse_args()
    logging.basicConfig(
        format="%(levelname)s: %(message)s",
        level=logging.DEBUG if args.debug else logging.INFO,
    )
    config = configparser.ConfigParser()
    config.read_file(args.configuration)
    cache = msal.SerializableTokenCache()

    if os.path.exists("%s/o365-imap-proxy-cache.bin" % XDG_CACHE_HOME):
        cache.deserialize(
            open("%s/o365-imap-proxy-cache.bin" % XDG_CACHE_HOME, "r").read()
        )
    else:
        logging.info(
            "No cache found at %s/o365-imap-proxy-cache.bin. This will be created once you've registered."
            % XDG_CACHE_HOME
        )

    atexit.register(
        lambda: open("%s/o365-imap-proxy-cache.bin" % XDG_CACHE_HOME, "w").write(
            cache.serialize()
        )
        if cache.has_state_changed
        else None
    )

    app = msal.PublicClientApplication(
        config["o365"]["client_id"],
        authority=config["o365"].get(
            "authority", "https://login.microsoftonline.com/common/"
        ),
        token_cache=cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(
            ["https://outlook.office.com/IMAP.AccessAsUser.All"], account=accounts[0]
        )

    if not result or "access_token" not in result:
        if args.register:
            flow = app.initiate_device_flow(
                scopes=["https://outlook.office.com/IMAP.AccessAsUser.All"]
            )
            if "user_code" not in flow:
                raise ValueError(
                    "Fail to create device flow: %s"
                    % json.dumps(flow, indent=4, sort_keys=True)
                )
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
            sys.exit(0)
        else:
            print("Please re-run this program with --register to set up your account.")
            sys.exit(1)

    username = accounts[0]["username"]
    auth_string = "user=%s\1auth=Bearer %s\1\1" % (username, result["access_token"])

    socketserver.TCPServer.allow_reuse_address = True
    logging.info("Ready for IMAP connections!")
    with socketserver.TCPServer(("127.0.0.1", 9999), ImapHandler) as server:
        server.serve_forever()
