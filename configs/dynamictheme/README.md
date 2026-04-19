# Dynamic Theme Configuration

Configuration for Dynamic Theme by Christophe Lavalle.

## Files:
- **ThemePersonalize.reg** - Windows personalization registry settings (dark mode)

## Note on settings.dat:
The settings.dat file (app-specific configuration) is locked by the Windows Store app runtime and cannot be exported while the system is running.

**Deployment Strategy:**
1. Install Dynamic Theme from Microsoft Store
2. Apply registry settings for dark mode
3. App will create default settings on first launch
4. Users can configure preferences in the app

## Settings Applied by Registry:
- Dark mode enabled for apps (AppsUseLightTheme = 0)
- Dark mode enabled for system (SystemUsesLightTheme = 0)
- Transparency enabled (EnableTransparency = 1)

## Manual Configuration:
After deployment, users can open Dynamic Theme and configure:
- Automatic switching schedule
- Wallpaper rotation preferences
- Theme transition settings

