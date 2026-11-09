# HQA Spring Staff Time Off

Staff request time off on a web page. HR and the Principal approve, ask for more info, or deny from a dashboard that shows who else is out, days used this school year, and busy weeks. admin@ can view the dashboard but can't decide. Everything is saved to a Google Sheet and emails go out automatically.

| File | What it is |
|---|---|
| `Code.gs`, `Logo.gs`, `Request.html`, `Review.html`, `Styles.html`, `appsscript.json` | The app: staff request page + HR dashboard |
| `time-off-trial.html` | Click-through demo with sample data. Download it and open it in a browser |

> GitHub only stores the code. The app runs in Google Apps Script. Keep this repo **private**.

## Set up

1. Make a new Google Sheet, then go to **Extensions > Apps Script**.
2. Create a file for each app file above (same names) and paste the code in. Skip the trial page.
   To edit `appsscript.json`, turn on **Project Settings > Show "appsscript.json"**.
3. Check `CONFIG` at the top of `Code.gs`.
4. Run `setup()` and accept the permissions. Optional: run `loadSampleRequests()` for example data.
5. **Deploy > New deployment > Web app**. Execute as: *Me*. Who has access: *Anyone within hqaspring.org*.

- Staff: the web app URL
- HR, Principal, admin: the web app URL + `?page=review`

To update the code later, use **Deploy > Manage deployments > Edit > New version** so the link stays the same.

## Who can do what

| | Request time off | See the dashboard | Approve / deny |
|---|---|---|---|
| Staff | ✓ (own requests only) | | |
| `APPROVERS` (principal@, hr@) | ✓ | ✓ | ✓ |
| `VIEWERS` (admin@) | ✓ | ✓ view only | |

## Emails

| When | Who gets it |
|---|---|
| New or resubmitted request | `APPROVERS` (cc `ADMIN_CC`); staff member gets a receipt |
| Approved / Needs Info / Denied | Staff member, with the note; FYI copy to `ADMIN_CC` and the other approver |
| Approved | `COVERAGE_NOTIFY` (secretary) for coverage. The reason is left out |
| Cancelled by staff | `APPROVERS` |
| Weekday mornings | `APPROVERS` get a summary of pending requests |

## Trial mode

`TEST_MODE: true` sends every email only to `TEST_EMAIL`, with a yellow bar showing who it would have gone to. Set `TEST_ROLE: 'viewer'` to see the admin (view only) dashboard.

## Going live

- Set `TEST_MODE: false`.
- Fill in `APPROVERS`, `VIEWERS`, `ADMIN_CC` and `COVERAGE_NOTIFY`.
- Check `LEAVE_TYPES`, `MIN_DAYS_NOTICE`, `ANNUAL_DAYS`, `NO_SCHOOL` and `BUSY` dates.
- Delete any sample rows from the **Time Off** tab.
- Own the Sheet and script from a school role account (for example hr@), and share the Sheet only with HR, the Principal and admin.

## Access and security

- Data lives in a Google Sheet in HQA's Google Workspace.
- Only hqaspring.org accounts can open the app. Sign-in is Google's, so there are no extra passwords.
- Permissions are checked on Google's server, not in the browser.
- The optional "Reason" field is visible only to the staff member, HR, the Principal and admin.
- Every submission and decision is logged with who and when.
