/**
 * ============================================================================
 * Palia Fish Assistant
 * Settings.gs
 *
 * Centralised settings and user preferences.
 * ============================================================================
 */

const APP_SETTINGS = Object.freeze({

  APP_NAME: "Palia Fish Assistant",

  DEFAULT_THEME: "pawtastic",

  HIGHLIGHT_COLOUR: "#FFF59D",

  REMEMBER_LAST_FISH: true,

  AUTO_HIGHLIGHT: true,

  SEARCH_CASE_SENSITIVE: false

});


/**
 * Returns all user settings.
 */
function getUserSettings() {

  const props = PropertiesService
    .getUserProperties();

  return {

    theme:
      props.getProperty(USER_PROPERTIES.THEME) ||
      APP_SETTINGS.DEFAULT_THEME,

    autoHighlight:
      readBoolean(
        props.getProperty(USER_PROPERTIES.AUTO_HIGHLIGHT),
        APP_SETTINGS.AUTO_HIGHLIGHT
      ),

    rememberFish:
      readBoolean(
        props.getProperty(USER_PROPERTIES.REMEMBER_FISH),
        APP_SETTINGS.REMEMBER_LAST_FISH
      ),

    lastFish:
      props.getProperty(USER_PROPERTIES.LAST_FISH) || "",

    lastRecipe:
      props.getProperty(USER_PROPERTIES.LAST_RECIPE) || ""

  };

}


/**
 * Saves a setting.
 */
function saveUserSetting(key, value) {

  const validKeys = Object.values(USER_PROPERTIES);

  if (!validKeys.includes(key)) {
    throw new Error("Unknown setting: " + key);
  }

  PropertiesService
    .getUserProperties()
    .setProperty(key, String(value));

}


/**
 * Saves the last lookup.
 */
function saveLastLookup(result) {

  if (!result) {
    return;
  }

  const props = PropertiesService.getUserProperties();

  props.setProperty(
    USER_PROPERTIES.LAST_FISH,
    result.fish
  );

  props.setProperty(
    USER_PROPERTIES.LAST_RECIPE,
    result.recipe
  );

}


/**
 * Clears all stored settings.
 */
function resetUserSettings() {

  PropertiesService
    .getUserProperties()
    .deleteAllProperties();

  saveUserSetting(
    USER_PROPERTIES.THEME,
    APP_SETTINGS.DEFAULT_THEME
  );

  saveUserSetting(
    USER_PROPERTIES.AUTO_HIGHLIGHT,
    APP_SETTINGS.AUTO_HIGHLIGHT
  );

  saveUserSetting(
    USER_PROPERTIES.REMEMBER_FISH,
    APP_SETTINGS.REMEMBER_LAST_FISH
  );

}


/**
 * Reads a boolean value safely.
 */
function readBoolean(value, defaultValue) {

  if (value === null) {
    return defaultValue;
  }

  return value === "true";

}

const USER_PROPERTIES = Object.freeze({

  THEME: "theme",

  AUTO_HIGHLIGHT: "autoHighlight",

  REMEMBER_FISH: "rememberFish",

  LAST_FISH: "lastFish",

  LAST_RECIPE: "lastRecipe"

});
