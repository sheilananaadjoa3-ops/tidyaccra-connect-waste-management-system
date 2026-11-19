# TidyAccra Connect: Waste Management System

A Visual Basic (VB.NET) Windows Forms app that connects residents in Accra, Ghana with nearby waste collectors. A resident requests a pickup, and the nearest online collector accepts it.

Student project (IT, UPSA). Prices, distances and payments are illustrative.

## Features
- **Account gate:** nobody can use the app without an account. Residents and collectors sign up; the admin uses a special email.
- **Google and Facebook sign-in:** a demo of the steps (not connected to the real services).
- **Forgot your password:** reset with the email and phone number used at sign-up.
- **Residents:** request a pickup, save an address, follow the live status, optional subscription (Basic, Standard, Business).
- **Payment (simulated):** a pickup without a plan asks to confirm a GH¢15 (estimate) mobile money payment.
- **Collectors:** go online, see requests within 3 km (nearest first), accept the job, mark it collected.
- **Admin:** dashboard of requests, collectors and residents; can widen the search to 5 km and assign only a collector inside the radius.
- **Saved data:** accounts, plans and pickups are saved to a file and loaded again at the next start.
- **No form designer needed:** every screen is built by code.

## How to run
1. Install Visual Studio with the ".NET desktop development" workload.
2. Open the project (the `.vbproj` file).
3. Set the Startup form to `frmWelcome` (the file `Form1.vb`) and press F5.

## Demo accounts
| Role | Email | Password |
|---|---|---|
| Admin | admin@tidyaccra.com | admin123 |
| Collector | kofi@example.com | pass123 |
| Collector | akua@example.com | pass123 |
| Resident (Standard plan) | ama@example.com | pass123 |
| Resident (no plan) | yaw@example.com | pass123 |

These are demo accounts only. Change the admin password in `modLink.vb` before any real use.

## Files
| File | Page / class | What it does |
|---|---|---|
| `Form1.vb` | `frmWelcome` | Log in, Create account, Google/Facebook (demo), Forgot password |
| `Form2.vb` | `frmSubscribe` | Optional subscription plans (Basic, Standard, Business) |
| `Form3.vb` | `frmResident` | Request a pickup, live status, payment |
| `Form4.vb` | `frmHome` | Home page |
| `Form5.vb` | `frmCollector` | Go online, accept a job, mark it collected |
| `Form6.vb` | `frmAdmin` | Admin dashboard: requests, collectors, residents |
| `modLink.vb` | shared module | Shared data, colours, matching, navigation bar, full screen, background |
| `modData.vb` | shared module | Saves and loads the data file |

Every page builds its own buttons, boxes and tables in code, so the `.Designer.vb` and `.resx` files stay empty.

## Limits
- Everything runs in one program on one computer. Several phones or computers would need a shared online database.
- Data is stored in a file, not a SQL server.
- Payment is simulated and Google/Facebook sign-in is a demo.
