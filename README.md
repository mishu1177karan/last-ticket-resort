# Last Ticket Resort

An OTC-style ticket desk. Buyers send how many tickets they want and a price range. Sellers send how many they have and an asking price. The public board shows demand and supply per event without names. The desk owner matches quotes, coordinates the handover in person, and charges a flat fee per side on a completed deal.

## Pieces

| Part | Where | What it does |
|---|---|---|
| `index.html` | GitHub Pages | Public board and quote form |
| `admin.html` | GitHub Pages | Matching desk (needs the admin key) |
| `config.js` | GitHub Pages | Holds the Apps Script web app URL |
| `apps-script/Code.gs` | Google Sheet → Extensions → Apps Script | API that reads and writes the Sheet |
| Google Sheet "Last Ticket Resort – Order Book" | Google Drive | `Quotes`, `Matches`, `Settings` tabs |

## How matching works

For each event, buyers are sorted by their max price (highest first) and sellers by their ask (lowest first). They pair up while the buyer's max is at least the seller's ask. The desk suggests one clearing price for the batch: the midpoint between the highest matched ask and the lowest matched buyer max. You can change it before confirming.

## Changing things

- **Fee:** change it on the desk page, or edit `FEE` in the Settings tab.
- **Admin key:** edit `ADMIN_KEY` in the Settings tab.
- **Backend code:** after editing `Code.gs` in Apps Script, go to Deploy → Manage deployments → edit → Version: New version → Deploy. The URL stays the same.
