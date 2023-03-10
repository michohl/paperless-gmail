#!/usr/bin/env python

import yaml
import pdfkit
from imap_tools import MailBox, AND, NOT, MailMessage, Header

def load_config(file_path: str) -> dict:
    with open(file_path) as f:
        return yaml.safe_load(f)

def format_filename(message: MailMessage) -> str:
    substitutions = {
        " ": "_",
        "/": "-",
        "!": "",
        "?": "",
        "@": "-",
        "Re: ": "",
        ":": ""
    }
    
    filename = f"{message.subject}--{message.date_str}"
    for char, replacement in substitutions.items():
        filename = filename.replace(char, replacement)

    return filename

# We have a way to "retrace" the thread but because Gmail actually quotes all the previous messages
# in each message we don't have to build it. We just need to find the last message and save that one instead
def rebuild_thread(mailbox: MailBox, threads: list, parent_message: MailMessage) -> str:
    html_compatibility = '<meta http-equiv="Content-type" content="text/html; charset=utf-8"/>'
    msg_id = threads.get(parent_message.headers["message-id"])
    while msg_id:
        headers=Header(name="message-id", value=msg_id[0])
       
        # This should only return one message but fetch returns a generator so we need to "loop"
        for msg in mailbox.fetch(AND(header=headers)):
            html = html_compatibility +  msg.html
            msg_id = threads.get(msg.headers["message-id"])
    return html

def main():
    config = load_config("settings.yaml")
    threads = dict()

    with MailBox(config["gmail"]["imap_server"]).login(config["gmail"]["email"], config["gmail"]["password"]) as mailbox:
        # Go through the high level labels defined by the user
        for mbox in config["fetch"]["mailboxes"]:

            # If there are children labels we should check those as well
            for label in mailbox.folder.list(mbox):
                uids = list()

                # Search for the current mailbox only
                mailbox.folder.set(label.name)

                # Return messages that match our custom label for ingesting and haven't been processed already
                for msg in mailbox.fetch(AND(NOT(gmail_label=config["fetch"]["consumed_label"]), gmail_label=config["fetch"]["label"])):
                    uids.append(msg.uid)

                    # prepend the HTML with extra compatibility
                    html = '<meta http-equiv="Content-type" content="text/html; charset=utf-8"/>' + msg.html

                    # output which messages we matched on
                    print(msg.subject)

                    # If we find a references section in the header then this
                    # message is a child in a thread/conversation that needs rebuilt
                    if msg.headers.get("references"):
                       #threads[msg.headers["message-id"]] = msg.headers["in-reply-to"]
                       threads[msg.headers["in-reply-to"]] = msg.headers["message-id"]
                       continue
                    elif msg.headers["message-id"] in threads.keys():
                        # If our message is not in the middle of a thread/conversation and our message-id
                        # appears in our threads dictionary that means this is the root node and we can use it
                        # to reverse lookup the last message in the thread which contains all the previous thread information
                        html = rebuild_thread(mailbox, threads, msg)
                        

                    # if the file has attachements we should pull a copy of them
                    for attachment in msg.attachments:
                        if attachment.filename.split(".")[-1].lower() in config["fetch"]["valid_extensions"]:
                            with open(f"{config['output']['paperless_directory']}/{attachment.filename}", "wb") as f:
                                f.write(attachment.payload)

                    # If we use external CSS for rendering it _will_ work but wkhtmltopdf will report an error
                    # but the error isn't actually stopping the PDF from rendering so we are throwing it to the void
                    formatted_filename = format_filename(msg)
                    output_location = f"{config['output']['paperless_directory']}/{formatted_filename}.pdf"
                    try:
                        pdfkit.from_string(
                            str(html), 
                            output_path=output_location, 
                            #options={"enable-local-file-access": None},
                        )
                    except:
                        continue
                
                # Mark consumed messages so they don't get duplicated repeatedly
                mailbox.copy(uids, config['fetch']['consumed_label'])

if __name__ == "__main__":
    main()
