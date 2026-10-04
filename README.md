# CustomerDay2026

A raffle wheel for Customer Day 2026. Upload the guest list, watch everyone's name fly onto the wheel, then click to spin.

## Running it

Open `index.html` in a browser (Edge or Chrome). No install or server needed, and it works offline.

## Guest list format

An Excel file (`.xlsx`, `.xls` or `.csv`) with a table like this:

| contact_name | account_name | invitation_accepted | raffle | privacy policy | photos |
|---|---|---|---|---|---|
| Mohammed Al Radi | Alfanar - London | Accepted | Y | Y | Y |

- Only rows with **raffle = Y** go into the draw. The bottom-left corner shows how many were excluded.
- `contact_name` and `raffle` are required; `account_name` is shown under the winner's name.
- The table can be on any sheet and doesn't need to start in cell A1.

## During the raffle

- Click the wheel, the SPIN button, or press **Space** to spin.
- When the winner is announced, click **Close** (or press **Enter**). Their name stays at the bottom of the screen.
- If the winner isn't in the room, click **Not here? Remove & redraw** to take them off the wheel and spin again.
- Press **F11** for full screen. Use the sound button to mute the ticks and fanfare.
