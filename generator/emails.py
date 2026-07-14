# -*- coding: utf-8 -*-
"""メール下書き(.eml)の作成。

X-Unsent ヘッダーを付けているため、Outlook でダブルクリックすると
「未送信メール」として開き、そのまま編集・送信できる。
"""
from __future__ import annotations

import mimetypes
from email.message import EmailMessage
from email import policy
from pathlib import Path


def build_eml(to: str, subject: str, body: str,
              attachments: list[Path], out: Path,
              bcc: list[str] | None = None,
              cc: list[str] | None = None) -> Path:
    msg = EmailMessage(policy=policy.SMTP)
    msg["To"] = to
    if cc:
        msg["Cc"] = ", ".join(cc)
    if bcc:
        msg["Bcc"] = ", ".join(bcc)
    msg["Subject"] = subject
    msg["X-Unsent"] = "1"
    msg.set_content(body, charset="utf-8")
    for path in attachments:
        path = Path(path)
        ctype, _ = mimetypes.guess_type(path.name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(
            path.read_bytes(), maintype=maintype, subtype=subtype,
            filename=path.name)
    out.write_bytes(msg.as_bytes())
    return out
