# DomainZax Sales Agent v1

Google Sheets + Apps Script email outreach workflow. This repository contains a safe-by-default starter implementation; it does not connect to your Google account or send anything until you deploy it and explicitly enable sending.

## Files
- `Code.gs`: spreadsheet setup, candidate domain seeding, draft generation, gated Gmail sending, follow-up logic, suppression list, activity log.
- `Domains.csv`: 19 owned domains are listed; `Domainzax.com` is retained as the brand and marked `NOT_FOR_SALE`. The other 18 are sale candidates but remain `REVIEW_REQUIRED` until each domain's sale status, asking/floor prices, and transfer route are confirmed.
- `Prospects-template.csv`: schema for manually verified prospects.

## Setup
1. Create a new Google Sheet owned by the account that will send mail.
2. Open **Extensions → Apps Script**.
3. Replace the editor contents with `Code.gs` from this folder and save.
4. Reload the spreadsheet; use **DomainZax Agent → 1. Set up / repair sheets**.
5. Use **DomainZax Agent → 2. Seed candidate domain list**. The 18 sale candidates are seeded with `owned_verified=YES` based on the owner's confirmation and `status=REVIEW_REQUIRED`. `Domainzax.com` is retained as the brand, set to `NOT_FOR_SALE`, and explicitly excluded from sales outreach.
6. Review the `Domains` sheet against your registrar accounts. For any of the 18 sale candidates, only set `status=ACTIVE` after confirming current ownership and sale availability; add an asking price, a minimum price, and a valid sales/marketplace URL. Do not change `Domainzax.com` from `NOT_FOR_SALE`.
7. Add prospects to `Prospects`. Required before draft creation: unique `prospect_id`, company, valid business email, public source URL, evidence supporting relevance, domain_match, relevance_score >= 80, `compliance_checked=YES`, and `status=READY`.
8. Run **Generate email drafts**. Inspect each draft in `Outreach`. No email is sent at this stage.
9. Before sending, verify the Gmail sender identity and that the public contact address/reply-to address actually works. Set the `REPLY_TO` and `PUBLIC_CONTACT_EMAIL` config values accordingly.
10. Check applicable marketing/privacy laws for each target market and document the basis for using the contact data. The `compliance_checked` field is an internal checkpoint, not legal advice.
11. Only after all checks, set `SEND_ENABLED=TRUE` in `Config`. Then install the daily trigger. You can stop sending immediately using **Disable all sending** or set `SEND_ENABLED=FALSE`.

## Safety defaults and limits
- Sending is disabled by default.
- First-contact cap: 5/day, shared daily cap for follow-ups is also applied by the current script.
- Relevance threshold: 80/100.
- Maximum two follow-ups, scheduled after 5 and 10 business days.
- Suppressed addresses are blocked.
- Domain ownership is checked against `owned_verified=YES` and `status=ACTIVE`.
- Prospects need a source URL and evidence; no guessed email addresses are generated.
- LinkedIn automation is intentionally excluded; use manually reviewed drafts because automated scraping/messaging may violate platform rules.
- The system uses built-in Google Apps Script services only. Google quotas apply and can change.

## Important limitations before production
- This is a v1 starter and should be tested with your own address first.
- Gmail reply detection depends on Gmail threading and sender identity; monitor the first runs and stop the trigger if replies are not detected correctly.
- Email deliverability, SPF/DKIM/DMARC, mailbox provider terms, legal basis, opt-out processing, and sender reputation must be checked before any cold outreach.
- Never market any candidate domain unless you have verified ownership, current sale status, asking price, and transfer route.
