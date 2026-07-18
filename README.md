# Favorik add-in — static shell

This repo hosts only the static presentation shell for the [Favorik](https://github.com/iprosites/favorik-2026)
Outlook add-in: the manifest, icons, and a small amount of HTML/JS that boots up and then talks
to **Favorik Companion**, a locally-run application on the user's own machine.

**No user data, files, or credentials pass through this repo or its hosted site.** Everything
here is published only because Outlook's sideload/installation process needs to fetch these
specific files (icons, manifest) from a publicly reachable server as part of validating the
add-in — it can't reach `localhost`. Once installed, the add-in's actual functionality (saving
attachments, opening folders, notes) runs entirely through the user's own local Companion
instance and never touches this site again for anything but re-loading its own UI shell.

Source and full documentation: https://github.com/iprosites/favorik-2026 (private).
