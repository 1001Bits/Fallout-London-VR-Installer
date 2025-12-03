# Fallout: London VR Installer

An automated installer for setting up Fallout: London to work with Fallout 4 VR.

## Features

- **Automatic Detection**: Finds your Fallout 4, Fallout 4 VR, GOG, and existing installations via registry and drive scanning
- **Smart DPI Scaling**: Adapts to different screen resolutions and Windows scaling settings
- **Update Mode**: Detects existing installations and offers streamlined updates
- **Fresh Install Mode**: Complete automated setup from scratch

## Fresh Installation Steps

The installer performs the following steps for a new installation:

1. **Download & Install Mod Organizer 2** - Downloads MO2 Portable and extracts it to your chosen location
2. **Download & Install F4SEVR** - Fallout 4 Script Extender for VR, installed to your Fallout 4 VR directory
3. **Download & Install FRIK** - Full Room-scale Immersive Kinematic body mod for VR
4. **Download & Install Comfort Swim VR** - Swimming comfort mod for VR
5. **Download & Install Buffout 4 NG** - Crash logging and stability improvements
6. **Copy MO2 Assets** - Pre-configured mod profiles and settings
7. **Copy Fallout Data with Smart DLC Detection** - Copies DLC files from Fallout 4, detecting pre-NG vs Next-Gen versions
8. **Copy FRIK Configuration** - Optimized FRIK settings
9. **Install xSE Plugin Preloader** - Required for F4SE plugins to load correctly
10. **Downgrade Next-Gen Files** (if needed) - Automatically downgrades Next-Gen BA2 archives to work with VR
11. **Create Desktop Shortcut** - Quick launch shortcut
12. **Create Start Menu Shortcuts** - Start menu folder with shortcuts for the game and Mod Organizer

## Update Mode Steps

When updating an existing installation:

1. **Backup INI Files** - Saves your fallout4.ini, fallout4prefs.ini, fallout4custom.ini
2. **Remove Deprecated Mods** - Cleans up old/incompatible mods:
   - Old FRIK versions (FRIK.v74, v75, v76, etc.)
   - High FPS Physics Fix
   - XDI mod directory
   - PrivateProfileRedirector F4
   - Version Check Patcher
   - Old F4SE folder organization (pre-0.96)
3. **Remove Old Fallout London VR Mod** - Ensures clean update
4. **Download & Install Latest FRIK** - Always updates to newest version
5. **Copy Updated Mod Files** - Merges new files while preserving your customizations
6. **Merge Mod List** - Adds new mods while preserving your custom mod order
7. **Upgrade to London 1.03** (optional) - If 1.03 files are detected
8. **Update INI Settings** - Applies latest recommended VR UI settings
9. **Install xSE Plugin Preloader** - Ensures latest version

## Requirements

- Fallout 4 VR (Steam or GOG)
- Fallout 4 with all DLC (for DLC assets)
- Fallout: London game files (downloaded separately from GOG)
- ~53 GB free disk space

## License

MIT License - see [LICENSE](LICENSE) for details.
