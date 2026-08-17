# Audio / Bluetooth connect logs

Connect, isolate, and disconnect write persistent logs here so a work machine can `git add` / `git push` them and this machine can `git pull` for debugging.

## Files

| File                                   | Role                                     |
| -------------------------------------- | ---------------------------------------- |
| `audio-bt-connect.log`                 | Append-only index (one line per attempt) |
| `audio-bt-connect-YYYYMMDD-HHmmss.log` | Full session block for that attempt      |
| `audio-bt-connect-latest.log`          | Copy of the most recent session          |

These files are **not** gitignored. After a failed connect at work:

```
git add docs/debug/audio-bt-connect*.log
git commit -m "AudioBt connect log from work"
git push
```

Then pull here and open `audio-bt-connect-latest.log`.
