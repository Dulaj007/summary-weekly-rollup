"""Decide whether a given Outlook MailItem belongs to a configured process."""


def is_process_email(mail_item, process_config: dict) -> bool:
    body = (getattr(mail_item, "Body", "") or "").lower()
    subject = (getattr(mail_item, "Subject", "") or "").lower()
    sender = (getattr(mail_item, "SenderEmailAddress", "") or "").lower()

    match_text = (process_config.get("match_text") or "").lower()
    if match_text and match_text not in body:
        return False

    subject_contains = process_config.get("subject_contains")
    if subject_contains and subject_contains.lower() not in subject:
        return False

    sender_contains = process_config.get("sender_contains")
    if sender_contains and sender_contains.lower() not in sender:
        return False

    return True
