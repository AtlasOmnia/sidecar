// Synchronous, tiny head entry: resolve stored/default scheme before styles/body paint.
// Office may not be available yet; the normal pane refreshes when Office.onReady fires.
import { applyTheme, browserStorage, colorSchemeMedia, currentOfficeTheme, loadTheme } from "./omniaTheme";
applyTheme(document.documentElement, loadTheme(browserStorage()), currentOfficeTheme(), colorSchemeMedia()?.matches ?? false);
