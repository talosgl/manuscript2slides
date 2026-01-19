## Troubleshooting

### Windows SmartScreen Warning

When launching the Windows binary for the first time, you may see a SmartScreen warning:

<img src="assets/imgs/win_smart_screen_popup.png" alt="Windows SmartScreen popup" width="500">

This is normal for unsigned applications. Click "More info":

<img src="assets/imgs/win_smart_screen_more_info.png" alt="Windows SmartScreen more info" width="500">

Then click "Run anyway" to launch the application.

### GUI won't launch

- Check log files in `~/Documents/manuscript2slides/logs/`
- On Linux, ensure Qt dependencies are installed (see above)

### Conversion fails

- Verify input file isn't corrupted (try opening in Word/PowerPoint)
- Check logs for specific error messages
- Try disabling "experimental formatting" if crashes occur

### Output looks wrong

- Some complex formatting may not transfer perfectly
- Try adjusting chunking strategy
- Check if your template has been modified significantly