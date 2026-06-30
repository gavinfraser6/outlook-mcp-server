"""Shared pytest fixtures and lightweight fakes for the Outlook MCP server.

These fakes mimic just enough of the Outlook COM surface (MailItem, Folder,
Items, Recipients, Attachments, Namespace, Application) to exercise the server
logic without a real mailbox. No live Outlook or credentials required.
"""

from __future__ import annotations

import datetime
import os
import sys
from typing import Any, Dict, List, Optional, Tuple

import pytest

# Make the project root importable when tests run from anywhere.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import outlook_helpers as H  # noqa: E402


# ---------------------------------------------------------------------------
# COM-shaped fakes
# ---------------------------------------------------------------------------

class FakeRecipient:
    def __init__(self, name: str, address: str):
        self.Name = name
        self.Address = address


class FakeRecipients:
    """Mimics the 1-indexed, callable COM Recipients collection."""

    def __init__(self, people: List[Tuple[str, str]]):
        self._people = people

    @property
    def Count(self) -> int:
        return len(self._people)

    def __call__(self, index: int) -> FakeRecipient:
        name, addr = self._people[index - 1]
        return FakeRecipient(name, addr)


class FakeAttachment:
    def __init__(self, filename: str, size: int, content: str = ""):
        self.FileName = filename
        self.Size = size
        self._content = content
        self.saved_to: Optional[str] = None

    def SaveAsFile(self, path: str) -> None:
        self.saved_to = path
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(self._content)


class FakeAttachments:
    def __init__(self, attachments: List[FakeAttachment]):
        self._items = attachments

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> FakeAttachment:
        return self._items[index - 1]


class FakeMail:
    """A configurable fake MailItem usable for both reading and writing."""

    def __init__(
        self,
        entry_id: str = "ID1",
        subject: str = "Subject",
        sender_name: str = "Alice Example",
        sender_email: str = "alice@example.com",
        body: str = "Hello world.",
        html_body: str = "",
        received: Optional[datetime.datetime] = None,
        sent: Optional[datetime.datetime] = None,
        recipients: Optional[List[Tuple[str, str]]] = None,
        attachments: Optional[List[FakeAttachment]] = None,
        unread: bool = True,
        importance: int = 1,
        flag_status: int = 0,
        categories: str = "",
        conversation_id: str = "CONV1",
    ):
        self.EntryID = entry_id
        self.Subject = subject
        self.SenderName = sender_name
        self.SenderEmailAddress = sender_email
        self.Body = body
        self.HTMLBody = html_body
        self.ReceivedTime = received if received is not None else datetime.datetime.now()
        self.SentOn = sent
        self.Recipients = FakeRecipients(recipients or [("Bob", "bob@example.com")])
        self.Attachments = FakeAttachments(attachments or [])
        self.UnRead = unread
        self.Importance = importance
        self.FlagStatus = flag_status
        self.Categories = categories
        self.ConversationID = conversation_id
        # write-side / outbound props
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.Sent = False
        self.saved = False
        self.sent_flag = False
        self.deleted = False
        self.moved_to: Optional["FakeFolder"] = None

    # outbound behaviour ----------------------------------------------------
    def Save(self):
        self.saved = True
        # Mirror Outlook: saving resolves the To/CC/BCC strings into the
        # Recipients collection. Only overwrite when something was set, so a
        # reply/forward draft keeps its pre-populated recipients.
        people = []
        for field in (self.To, self.CC, self.BCC):
            for addr in H.parse_recipients(field):
                people.append(("", addr))
        if people:
            self.Recipients = FakeRecipients(people)

    def Send(self):
        self.sent_flag = True
        self.Sent = True

    def Delete(self):
        self.deleted = True

    def Move(self, folder):
        self.moved_to = folder

    def Reply(self):
        return FakeMail(entry_id="REPLY1", subject="RE: " + self.Subject,
                        body="\n\n> " + self.Body, conversation_id=self.ConversationID)

    def ReplyAll(self):
        r = self.Reply()
        r.EntryID = "REPLYALL1"
        return r

    def Forward(self):
        return FakeMail(entry_id="FWD1", subject="FW: " + self.Subject,
                        body="\n\n---- Forwarded ----\n" + self.Body,
                        attachments=list(self.Attachments._items),
                        conversation_id=self.ConversationID)


class FakeItems(list):
    """List subclass exposing the COM Items methods used by the server."""

    def Sort(self, *args, **kwargs):
        return None

    @property
    def Count(self) -> int:
        return len(self)

    def Restrict(self, query: str) -> "FakeItems":
        q = query.lower()
        if "[unread] = true" in q:
            return FakeItems([m for m in self if getattr(m, "UnRead", False)])
        return self


class FakeFolder:
    def __init__(self, name: str, mails: Optional[List[FakeMail]] = None):
        self.Name = name
        self._items = FakeItems(mails or [])
        self.Folders: List["FakeFolder"] = []

    @property
    def Items(self) -> FakeItems:
        return self._items


class FakeNamespace:
    def __init__(self, inbox: List[FakeMail]):
        self.folders_by_index = {
            3: FakeFolder("Deleted Items"),
            5: FakeFolder("Sent Items"),
            6: FakeFolder("Inbox", inbox),
            9: FakeFolder("Calendar"),
            13: FakeFolder("Tasks"),
            16: FakeFolder("Drafts"),
        }
        archive = FakeFolder("Archive")
        attended = FakeFolder("Attended")
        self.Folders = [self.folders_by_index[6], archive, attended]
        self.folders_by_index[6].Folders = [archive, attended]
        self._by_id: Dict[str, FakeMail] = {m.EntryID: m for m in inbox}
        self.Categories = [
            _cat("Work", 1), _cat("Personal", 2),
        ]

    def register(self, mail: FakeMail):
        self._by_id[mail.EntryID] = mail

    def GetDefaultFolder(self, index: int) -> FakeFolder:
        return self.folders_by_index[index]

    def GetItemFromID(self, entry_id: str) -> FakeMail:
        if entry_id not in self._by_id:
            raise KeyError(entry_id)
        return self._by_id[entry_id]


def _cat(name, color):
    class C:
        pass
    c = C()
    c.Name = name
    c.Color = color
    return c


class FakeOutlook:
    def __init__(self, namespace: FakeNamespace):
        self._ns = namespace
        self.created: List[FakeMail] = []

    def CreateItem(self, item_type: int) -> FakeMail:
        m = FakeMail(entry_id=f"NEW{len(self.created)+1}", subject="", body="")
        m.Recipients = FakeRecipients([])
        self.created.append(m)
        self._ns.register(m)
        return m

    def GetNamespace(self, _profile: str) -> FakeNamespace:
        return self._ns


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_inbox() -> List[FakeMail]:
    now = datetime.datetime.now()
    return [
        FakeMail(entry_id="A", subject="Invoice #42 due Friday",
                 sender_name="Acme Billing", sender_email="billing@acme.com",
                 body="Please pay invoice 42 by Friday.",
                 received=now - datetime.timedelta(hours=1),
                 attachments=[FakeAttachment("invoice.pdf", 1024)],
                 unread=True, importance=2, conversation_id="C1"),
        FakeMail(entry_id="B", subject="Lunch tomorrow?",
                 sender_name="Bob Friend", sender_email="bob@friends.com",
                 body="Want to grab lunch tomorrow?",
                 received=now - datetime.timedelta(hours=5),
                 unread=False, conversation_id="C2"),
        FakeMail(entry_id="C", subject="Project update notes",
                 sender_name="Carol PM", sender_email="carol@work.com",
                 body="Here are the notes. spam-word inside.",
                 received=now - datetime.timedelta(days=2),
                 categories="Work", unread=True, conversation_id="C3"),
    ]


@pytest.fixture
def server(monkeypatch, sample_inbox):
    """Import the server with COM connections monkeypatched to fakes."""
    import outlook_mcp_server as s

    ns = FakeNamespace(sample_inbox)
    outlook = FakeOutlook(ns)
    monkeypatch.setattr(s, "_WIN32_AVAILABLE", True)
    monkeypatch.setattr(s, "_connect", lambda: (outlook, ns))
    monkeypatch.setattr(s, "_namespace", lambda: ns)
    s._email_cache.clear()
    return s, ns, outlook
