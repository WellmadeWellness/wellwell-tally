# Setup – GitHub Pages + Firebase (about 15 minutes, free)

You need one Google account for Firebase and the GitHub repository with this code.

## 1. Create the Firebase project
1. Go to https://console.firebase.google.com and click **Create a project**. Name it e.g. `wellwell-tally`. Google Analytics is not needed.
2. Stay on the free **Spark** plan. No card is needed.

## 2. Turn on the database
1. Left menu: **Build → Firestore Database → Create database**.
2. Location: **asia-northeast3 (Seoul)**. Start in **production mode**.
3. Open the **Rules** tab, delete what is there, paste the whole content of `firestore.rules` from this repository, and click **Publish**.

## 3. Turn on sign-in
1. Left menu: **Build → Authentication → Get started**.
2. **Sign-in method** tab: enable **Google**, and also **Email/Password** if some people have no Google account.
3. **Settings** tab → **Authorized domains** → **Add domain**: `YOUR-GITHUB-NAME.github.io`.

## 4. Connect the page to Firebase
1. **Project settings** (gear icon) → **Your apps** → web icon `</>` → register an app named `wellwell-tally` (no Firebase Hosting).
2. Copy the `firebaseConfig` values into `docs/firebase-config.js` in this repository (replace the `YOUR_…` placeholders) and commit.

## 5. Make yourself the first admin
1. Open **Firestore Database → Data → Start collection**. Collection ID: `roles`.
2. Document ID: **your email in lower case**, e.g. `factory@wellwell.co.kr`.
3. Add a field `role` (string) = `admin`. Save.
After this, add everyone else from the page's **Team** tab.

## 6. Publish on GitHub Pages
1. In the GitHub repository: **Settings → Pages**.
2. Source: **Deploy from a branch** → branch `main`, folder `/docs` → **Save**.
3. After a minute the page is at `https://YOUR-GITHUB-NAME.github.io/REPOSITORY-NAME/`.

## 7. First use
1. Open the page and sign in with the admin email.
2. **Rules & prices → Import prices (.xlsx)** and load the company price list (columns `Product`, `Price`). It is stored in Firebase for everyone; the price list is never put in GitHub.
3. **Team** tab: add uploaders and viewers by email.
4. Upload this month's files. Every day, upload that day's files; pick an earlier month from the Month list to look back.

## Good to know
- Free limits (Spark plan): 1 GB stored, 50,000 reads and 20,000 writes per day – far above what this tool uses. Check the current figures on Firebase's pricing page.
- Customer names, phones and addresses in the 발주 files are read in the browser and never stored.
- At the end of each month, open the month, click **Download Excel**, and keep the file on the company drive as a backup.
