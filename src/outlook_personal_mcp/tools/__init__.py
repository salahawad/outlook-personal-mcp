from . import mail, folders, drafts, calendar


def register_all(mcp, client, settings):
    folders.register(mcp, client)
    mail.register(mcp, client, settings)
    drafts.register(mcp, client, settings)
    calendar.register(mcp, client)
