# Improvement Workflow

Each improvement for the Beta 7 milestone is applied directly to `ATP_Beta6_v5.py`. We reuse the same file so that all updates build on top of the previous work. After we finish implementing and testing an improvement, the change is committed, keeping the existing file name and history intact. This way you do **not** need to copy the code back into your own file after every step—the repository already contains the latest version.

## Getting the latest file to test

1. Pull the latest commit from the repository (or download the updated archive if you are not using Git). The `ATP_Beta6_v5.py` file in the repo root is always the current version.
   - If you are using Git on the command line, open a terminal in your local clone and run:

     ```bash
     git fetch origin
     git pull origin <branch-name>
     ```

     Replace `<branch-name>` with the branch you are tracking (for example, `main`). These commands download the newest commits from GitHub and merge them into your local copy.
   - If you prefer the GitHub web interface, navigate to the repository page, click **Code ▾**, and choose **Download ZIP** to grab the latest snapshot. Extract the archive and copy `ATP_Beta6_v5.py` into your testing environment.
2. Open or run that file directly in your environment—for example, `python ATP_Beta6_v5.py`—to exercise the newest changes.
3. If desired, make a backup copy before you test so you can compare against earlier versions.

If you prefer to keep personal backups, you can copy `ATP_Beta6_v5.py` elsewhere before we start, but it is optional. The assisted workflow updates the tracked file in-place and preserves the Beta 6 baseline in version control for reference.
