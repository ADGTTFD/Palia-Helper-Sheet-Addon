/**
 * ============================================================================
 * Palia Fish Assistant
 * Lookup.gs
 * ============================================================================
 *
 * Dynamic lookup engine.
 *
 * The spreadsheet is always treated as the source of truth.
 *
 * Features
 * --------
 * • Automatic workbook discovery
 * • Dynamic fish list
 * • Dynamic recipe list
 * • Automatic recipe variant detection
 * • Quantity calculations
 * • Cell highlighting
 * • Sidebar API
 *
 * No fish names or recipe names are hardcoded.
 * ============================================================================
 */


/**
 * --------------------------------------------------------------------------
 * Spreadsheet Configuration
 * --------------------------------------------------------------------------
 */

const LOOKUP_CONFIG = Object.freeze({

  HEADER_ROW: 1,

  FIRST_DATA_ROW: 2,

  FIRST_DATA_COLUMN: 1,

  DEFAULT_QUANTITY: 1

});


/**
 * --------------------------------------------------------------------------
 * Recipe Variant Suffixes
 *
 * Used when splitting headers like:
 *
 * Maws Pokebowl Dari
 * Maws Pokebowl Heat Root
 * Maws Pokebowl SL / Sprout
 * --------------------------------------------------------------------------
 */

const VARIANT_SUFFIXES = Object.freeze([

  "Dari",

  "Heat Root",

  "SL / Sprout",

  "Sprout / SL",

  "Sprout",

  "SL",

  "Star Quality",

  "Star"

]);


/**
 * --------------------------------------------------------------------------
 * Lookup Service
 *
 * Everything in this file hangs from this object.
 * --------------------------------------------------------------------------
 */

const LookupService = {};

/**
 * ============================================================================
 * WORKBOOK DISCOVERY
 * ============================================================================
 */

/**
 * --------------------------------------------------------------------------
 * Returns the workbook descriptor.
 * --------------------------------------------------------------------------
 */
LookupService.getWorkbook = function () {

  return LookupService.discoverWorkbook();

};


/**
 * --------------------------------------------------------------------------
 * Finds the worksheet containing the Fish Profit table.
 *
 * The first cell containing "Fish" is treated as the start of the table.
 * --------------------------------------------------------------------------
 */
LookupService.discoverWorkbook = function () {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheets =
    spreadsheet.getSheets();

  for (const sheet of sheets) {

    const descriptor =
      LookupService.inspectSheet(sheet);

    if (descriptor) {
      return descriptor;
    }

  }

  throw new Error(
    "Unable to locate the Fish Profit table."
  );

};


/**
 * --------------------------------------------------------------------------
 * Checks whether a worksheet contains the lookup table.
 * --------------------------------------------------------------------------
 */
LookupService.inspectSheet = function (sheet) {

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();

  if (values.length === 0) {
    return null;
  }

  for (let row = 0; row < values.length; row++) {

    for (let col = 0; col < values[row].length; col++) {

      if (
        LookupService.normalise(values[row][col]) ===
        "fish"
      ) {

        return LookupService.buildWorkbookDescriptor(
          sheet,
          values,
          row,
          col
        );

      }

    }

  }

  return null;

};


/**
 * --------------------------------------------------------------------------
 * Builds a descriptor describing the discovered lookup table.
 * --------------------------------------------------------------------------
 */
LookupService.buildWorkbookDescriptor = function (
  sheet,
  values,
  headerRow,
  fishColumn
) {

  const headers = [];

  for (

    let column = fishColumn + 1;

    column < values[headerRow].length;

    column++

  ) {

    const header =
      String(values[headerRow][column]).trim();

    if (header === "") {
      continue;
    }

    headers.push({

      column: column + 1,

      header: header

    });

  }

  return {

    spreadsheet:
      SpreadsheetApp.getActiveSpreadsheet(),

    sheet: sheet,

    headerRow: headerRow + 1,

    fishColumn: fishColumn + 1,

    headers: headers,

    lastRow: sheet.getLastRow(),

    lastColumn: sheet.getLastColumn()

  };

};


/**
 * --------------------------------------------------------------------------
 * Normalises text for comparisons.
 * --------------------------------------------------------------------------
 */
LookupService.normalise = function (text) {

  return String(text)

    .replace(/\r/g, "")

    .replace(/\n/g, " ")

    .replace(/\s+/g, " ")

    .trim()

    .toLowerCase();

};

/**
 * ============================================================================
 * FISH DISCOVERY
 * ============================================================================
 */

/**
 * --------------------------------------------------------------------------
 * Returns every fish in the workbook.
 *
 * Used by:
 *   • Sidebar searchable list
 *   • Validation
 * --------------------------------------------------------------------------
 */
LookupService.getFishList = function () {

  const workbook =
    LookupService.getWorkbook();

  return LookupService.readFishList(workbook);

};


/**
 * --------------------------------------------------------------------------
 * Reads every fish name from the workbook.
 * --------------------------------------------------------------------------
 */
LookupService.readFishList = function (workbook) {

  const firstRow =
    workbook.headerRow + 1;

  const rowCount =
    workbook.lastRow - workbook.headerRow;

  if (rowCount <= 0) {
    return [];
  }

  const values =
    workbook.sheet
      .getRange(
        firstRow,
        workbook.fishColumn,
        rowCount,
        1
      )
      .getDisplayValues();

  return values

    .map(row => String(row[0]).trim())

    .filter(name => name !== "")

    .sort((a, b) =>
      a.localeCompare(
        b,
        undefined,
        {
          sensitivity: "base"
        }
      )
    );

};


/**
 * --------------------------------------------------------------------------
 * Finds a fish inside the workbook.
 *
 * Returns:
 *
 * {
 *    rowIndex,
 *    row
 * }
 *
 * or null.
 * --------------------------------------------------------------------------
 */
LookupService.findFishRow = function (fishName) {

  const workbook =
    LookupService.getWorkbook();

  const firstRow =
    workbook.headerRow + 1;

  const rowCount =
    workbook.lastRow - workbook.headerRow;

  const columnCount =
    workbook.lastColumn;

  const rows =
    workbook.sheet
      .getRange(
        firstRow,
        1,
        rowCount,
        columnCount
      )
      .getValues();

  const target =
    LookupService.normalise(fishName);

  for (let i = 0; i < rows.length; i++) {

    const current =
      LookupService.normalise(
        rows[i][workbook.fishColumn - 1]
      );

    if (current === target) {

      return {

        rowIndex:
          firstRow + i,

        row:
          rows[i]

      };

    }

  }

  return null;

};


/**
 * --------------------------------------------------------------------------
 * Returns true if the fish exists.
 * --------------------------------------------------------------------------
 */
LookupService.hasFish = function (fishName) {

  return (
    LookupService.findFishRow(fishName)
    !== null
  );

};

/**
 * ============================================================================
 * RECIPE DISCOVERY
 * ============================================================================
 */

/**
 * --------------------------------------------------------------------------
 * Returns every recipe discovered in the workbook.
 *
 * Example:
 *
 * [
 *   "Maws Pokebowl",
 *   "Fish Stew",
 *   "Sushi"
 * ]
 * --------------------------------------------------------------------------
 */
LookupService.getRecipeList = function () {

  const workbook =
    LookupService.getWorkbook();

  const recipes =
    LookupService.discoverRecipes(workbook);

  return recipes.map(recipe => recipe.name);

};


/**
 * --------------------------------------------------------------------------
 * Discovers every recipe and groups its variants.
 * --------------------------------------------------------------------------
 */
LookupService.discoverRecipes = function (workbook) {

  const recipes = {};

  workbook.headers.forEach(header => {

    const parsed =
      LookupService.splitRecipeHeader(
        header.header
      );

    if (!recipes[parsed.recipe]) {

      recipes[parsed.recipe] = {

        name: parsed.recipe,

        columns: []

      };

    }

    recipes[parsed.recipe].columns.push({

      columnIndex: header.column,

      header: header.header,

      label: parsed.variant

    });

  });

  return Object.values(recipes)

    .sort((a, b) =>
      a.name.localeCompare(
        b.name,
        undefined,
        {
          sensitivity: "base"
        }
      )
    );

};


/**
 * --------------------------------------------------------------------------
 * Splits a worksheet header into:
 *
 * Recipe Name
 * Variant
 *
 * Example:
 *
 * "Maws Pokebowl Dari"
 *
 * becomes
 *
 * {
 *    recipe: "Maws Pokebowl",
 *    variant: "Dari"
 * }
 * --------------------------------------------------------------------------
 */
LookupService.splitRecipeHeader = function (header) {

  const text =
    String(header).trim();

  for (const suffix of VARIANT_SUFFIXES) {

    if (text.endsWith(suffix)) {

      return {

        recipe:
          text.substring(
            0,
            text.length - suffix.length
          ).trim(),

        variant:
          suffix

      };

    }

  }

  return {

    recipe: text,

    variant: "Standard"

  };

};


/**
 * --------------------------------------------------------------------------
 * Finds every worksheet column belonging to a recipe.
 *
 * Returns:
 *
 * [
 *   {
 *      columnIndex,
 *      header,
 *      label
 *   }
 * ]
 * --------------------------------------------------------------------------
 */
LookupService.findRecipeColumns = function (recipeName) {

  const workbook =
    LookupService.getWorkbook();

  const recipes =
    LookupService.discoverRecipes(workbook);

  const recipe =
    recipes.find(item =>

      LookupService.normalise(item.name) ===
      LookupService.normalise(recipeName)

    );

  if (!recipe) {
    return [];
  }

  return recipe.columns;

};


/**
 * --------------------------------------------------------------------------
 * Returns true if a recipe exists.
 * --------------------------------------------------------------------------
 */
LookupService.hasRecipe = function (recipeName) {

  return (
    LookupService.findRecipeColumns(recipeName)
      .length > 0
  );

};

/**
 * ============================================================================
 * Lookup Engine
 * ============================================================================
 */

/**
 * Finds a fish row.
 */
LookupService.findFishRow = function (fishName) {

  const rows = LookupService.getFishRows();

  const target = String(fishName)
    .trim()
    .toLowerCase();

  for (let i = 0; i < rows.length; i++) {

    const value = String(rows[i][0])
      .trim()
      .toLowerCase();

    if (value === target) {

      return {

        rowIndex:
          LOOKUP_CONFIG.FIRST_DATA_ROW + i,

        row:
          rows[i]

      };

    }

  }

  return null;

};


/**
 * Finds every spreadsheet column belonging
 * to the selected recipe.
 */
LookupService.findRecipeColumns = function (recipeName) {

  const headers =
    LookupService.getHeaders();

  const columns = [];

  headers.forEach(function (header, index) {

    if (!header) {
      return;
    }

    const info =
      LookupService.splitRecipeHeader(header);

    if (
      info.recipe.toLowerCase() ===
      recipeName.toLowerCase()
    ) {

      columns.push({

        columnIndex: index + 1,

        header: header,

        label: info.variant

      });

    }

  });

  return columns;

};


/**
 * Converts spreadsheet values into numbers.
 */
LookupService.toNumber = function (value) {

  if (
    value === null ||
    value === "" ||
    value === undefined
  ) {
    return 0;
  }

  if (typeof value === "number") {
    return value;
  }

  const cleaned = String(value)
    .replace(/,/g, "")
    .trim();

  const number = Number(cleaned);

  return Number.isNaN(number)
    ? 0
    : number;

};


/**
 * Performs a lookup.
 */
LookupService.lookupFishRecipe = function (request) {

  if (!request) {
    throw new Error("Lookup request missing.");
  }

  const fish =
    String(request.fish || "").trim();

  const recipe =
    String(request.recipe || "").trim();

  let quantity =
    Number(request.quantity);

  if (
    Number.isNaN(quantity) ||
    quantity < 1
  ) {
    quantity = 1;
  }

  const fishRow =
    LookupService.findFishRow(fish);

  if (!fishRow) {
    throw new Error(
      'Fish "' + fish + '" not found.'
    );
  }

  const recipeColumns =
    LookupService.findRecipeColumns(recipe);

  if (recipeColumns.length === 0) {
    throw new Error(
      'Recipe "' + recipe + '" not found.'
    );
  }

  const variants = [];

  recipeColumns.forEach(function (column) {

    const value =
      LookupService.toNumber(
        fishRow.row[column.columnIndex - 1]
      );

    variants.push({

      label: column.label,

      header: column.header,

      value: value,

      total: value * quantity,

      row: fishRow.rowIndex,

      column: column.columnIndex

    });

  });

  return {

    fish: fish,

    recipe: recipe,

    quantity: quantity,

    variants: variants

  };

};

/**
 * ============================================================================
 * Recipe Discovery
 * ============================================================================
 */

/**
 * Returns every recipe discovered in the worksheet.
 *
 * Output:
 *
 * [
 *   {
 *     name: "Poke Bowl",
 *     columns: [
 *       {
 *         column: 5,
 *         header: "Poke Bowl",
 *         variant: "Standard"
 *       },
 *       {
 *         column: 6,
 *         header: "Poke Bowl Dari",
 *         variant: "Dari"
 *       }
 *     ]
 *   }
 * ]
 */
LookupService.getRecipes = function () {

    const workbook =
        LookupService.getWorkbook();

    const recipes = {};

    workbook.headers.forEach(header => {

        const parsed =
            LookupService.parseRecipeHeader(
                header.header
            );

        if (!recipes[parsed.recipe]) {

            recipes[parsed.recipe] = {

                name: parsed.recipe,

                columns: []

            };

        }

        recipes[parsed.recipe].columns.push({

            column: header.column,

            header: header.header,

            variant: parsed.variant

        });

    });

    return Object.values(recipes)

        .sort((a, b) =>
            a.name.localeCompare(
                b.name,
                undefined,
                { sensitivity: "base" }
            )
        );

};

/**
 * Parses a recipe header into:
 *
 * Recipe Name
 * Variant
 *
 * Examples:
 *
 * Poke Bowl
 *      -> Standard
 *
 * Poke Bowl Dari
 *      -> Dari
 *
 * Poke Bowl Heat Root
 *      -> Heat Root
 */
LookupService.parseRecipeHeader = function (header) {

    const text =
        String(header).trim();

    const suffixes = [

        "Dari",

        "Heat Root",

        "SL / Sprout",

        "Sprout",

        "Star Quality",

        "Star"

    ];

    for (const suffix of suffixes) {

        if (text.endsWith(suffix)) {

            return {

                recipe:
                    text.substring(
                        0,
                        text.length - suffix.length
                    ).trim(),

                variant: suffix

            };

        }

    }

    return {

        recipe: text,

        variant: "Standard"

    };

};

/**
 * Returns every unique recipe name.
 *
 * Used by the sidebar search.
 */
LookupService.getRecipeList = function () {

    return LookupService

        .getRecipes()

        .map(recipe => recipe.name);

};

/**
 * ============================================================================
 * Lookup Engine
 * ============================================================================
 */

/**
 * Performs a lookup.
 *
 * Returns:
 *
 * {
 *   fish,
 *   recipe,
 *   quantity,
 *   variants:[]
 * }
 */
LookupService.lookupFishRecipe = function (request) {

    if (!request) {
        throw new Error("Lookup request is missing.");
    }

    const fish =
        String(request.fish || "").trim();

    const recipe =
        String(request.recipe || "").trim();

    let quantity =
        Number(request.quantity);

    if (
        Number.isNaN(quantity) ||
        quantity < 1
    ) {

        quantity = 1;

    }

    const workbook =
        LookupService.getWorkbook();

    const fishRow =
        LookupService.findFishRow(
            workbook,
            fish
        );

    if (!fishRow) {

        throw new Error(
            'Fish "' +
            fish +
            '" was not found.'
        );

    }

    const recipeInfo =
        LookupService.getRecipes()
            .find(r =>
                r.name.toLowerCase() ===
                recipe.toLowerCase()
            );

    if (!recipeInfo) {

        throw new Error(
            'Recipe "' +
            recipe +
            '" was not found.'
        );

    }

    const variants = [];

    recipeInfo.columns.forEach(column => {

        const value = workbook.sheet
            .getRange(
                fishRow.rowIndex,
                column.column
            )
            .getValue();

        const numeric =
            LookupService.toNumber(value);

        variants.push({

            label: column.variant,

            header: column.header,

            value: numeric,

            total: numeric * quantity,

            row: fishRow.rowIndex,

            column: column.column

        });

    });

    return {

        fish: fish,

        recipe: recipe,

        quantity: quantity,

        variants: variants

    };

};

/**
 * Converts spreadsheet values into numbers.
 */
LookupService.toNumber = function (value) {

    if (
        value === "" ||
        value === null ||
        value === undefined
    ) {

        return 0;

    }

    if (typeof value === "number") {
        return value;
    }

    const cleaned =
        String(value)
            .replace(/,/g, "")
            .trim();

    const number =
        Number(cleaned);

    return Number.isNaN(number)
        ? 0
        : number;

};

/**
 * Finds a fish within the workbook.
 */
LookupService.findFishRow = function (
    workbook,
    fishName
) {

    const target =
        fishName.toLowerCase();

    const rows =
        LookupService.getFishRows();

    for (let i = 0; i < rows.length; i++) {

        const name =
            String(rows[i][0]).trim().toLowerCase();

        if (name === target) {

            return {

                rowIndex:
                    workbook.headerRow + i + 1,

                row:
                    rows[i]

            };

        }

    }

    return null;

};

/**
 * ============================================================================
 * Public API
 *
 * These functions are called from Sidebar.html
 * using google.script.run().
 * ============================================================================
 */

/**
 * Returns every fish in the workbook.
 */
function getFishList() {

    return LookupService.getFishList();

}


/**
 * Returns every recipe discovered
 * from the workbook.
 */
function getRecipeList() {

    return LookupService.getRecipeList();

}


/**
 * Performs a lookup.
 */
function lookupFishRecipe(
    fish,
    recipe,
    quantity
) {

    const result =
        LookupService.lookupFishRecipe({

            fish: fish,

            recipe: recipe,

            quantity: quantity

        });

    const settings =
        getUserSettings();

    if (settings.autoHighlight) {

        highlightLookupResult(result);

    }

    if (settings.rememberFish) {

        saveLastLookup(result);

    }

    return result;

}


/**
 * Refresh hook.
 *
 * Reserved for future cache support.
 */
function refreshAssistantData() {

    SpreadsheetApp.flush();

    return {

        success: true,

        refreshed: new Date()

    };

}

/**
 * Refreshes runtime data.
 */
LookupService.refresh = function () {

    return refreshAssistantData();

};

/**
 * ============================================================================
 * Highlight Engine
 * ============================================================================
 */

/**
 * Highlights every matching result cell.
 */
function highlightLookupResult(result) {

    if (
        !result ||
        !result.variants ||
        result.variants.length === 0
    ) {
        return;
    }

    const workbook =
        LookupService.getWorkbook();

    clearLookupHighlights(workbook);

    const sheet =
        workbook.sheet;

    result.variants.forEach(function (variant) {

        sheet
            .getRange(
                variant.row,
                variant.column
            )
            .setBackground(
                APP_SETTINGS.HIGHLIGHT_COLOUR
            );

    });

    // Jump to the selected fish row
    sheet.setActiveRange(
        sheet.getRange(
            result.variants[0].row,
            workbook.fishColumn
        )
    );

}

/**
 * Removes all assistant highlights.
 */
function clearLookupHighlights(workbook) {

    const firstRow =
        workbook.headerRow + 1;

    const totalRows =
        workbook.lastRow -
        workbook.headerRow;

    if (totalRows <= 0) {
        return;
    }

    const firstRecipeColumn =
        workbook.fishColumn + 1;

    const recipeColumns =
        workbook.lastColumn -
        workbook.fishColumn;

    workbook.sheet
        .getRange(
            firstRow,
            firstRecipeColumn,
            totalRows,
            recipeColumns
        )
        .setBackground(null);

}

/**
 * Clears lookup highlighting.
 */
function clearLookupHighlight() {

    clearLookupHighlights(
        LookupService.getWorkbook()
    );

}

/**
 * Removes existing highlights.
 */
function clearHighlights() {

    clearLookupHighlight();

}
