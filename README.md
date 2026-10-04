# CustomerDay2026

A raffle wheel for Customer Day 2026. Upload the guest list, watch everyone's name fly onto the wheel, then click to spin.

## Running it

Open `index.html` in a browser (Edge or Chrome). No install or server needed, and it works offline.

## Guest list format

An Excel file (`.xlsx`, `.xls` or `.csv`) with a table like this:

| contact_name | account_name | invitation_accepted | raffle | privacy policy | photos |
|---|---|---|---|---|---|
| Mohammed Al Radi | Alfanar - London | Accepted | Y | Y | Y |

- Only rows with **raffle = Y** go into the draw. The header shows how many were excluded.
- `contact_name` and `raffle` are required; `account_name` is shown under the winner's name.
- The table can be on any sheet and doesn't need to start in cell A1.

## During the raffle

- Click the wheel, the SPIN button, or press **Space** to spin.
- After a winner is shown, choose **Remove from wheel** (press **Enter**) so they can't win twice, or **Keep on wheel**.
- Winners are listed on the right. Use the sound button to mute the ticks and fanfare.
