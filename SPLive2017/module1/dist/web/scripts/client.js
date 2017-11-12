(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else if(typeof exports === 'object')
		exports["module1"] = factory();
	else
		root["module1"] = factory();
})(this, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 1);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Class for managing Microsoft Teams themes
 * idea borrowed from the Dizz: https://github.com/richdizz/Microsoft-Teams-Tab-Themes/blob/master/app/config.html
 * Uses a hierarchical styles approach with a simple stylesheet
 */
var TeamsTheme = /** @class */ (function () {
    function TeamsTheme() {
    }
    /**
     * Set up themes on a page
     */
    TeamsTheme.fix = function (context) {
        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(TeamsTheme.themeChanged);
        if (context) {
            TeamsTheme.themeChanged(context.theme);
        }
        else {
            microsoftTeams.getContext(function (context) {
                TeamsTheme.themeChanged(context.theme);
            });
        }
    };
    /**
     * Manages theme changes
     * @param theme default|contrast|dark
     */
    TeamsTheme.themeChanged = function (theme) {
        var bodyElement = document.getElementsByTagName("body")[0];
        switch (theme) {
            case "dark":
            case "contrast":
                bodyElement.className = "theme-" + theme;
                break;
            case "default":
                bodyElement.className = "";
        }
    };
    return TeamsTheme;
}());
exports.TeamsTheme = TeamsTheme;


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__(2);


/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

// Default entry point for client scripts
// Automatically generated
// Please avoid from modifying to much...
function __export(m) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}
Object.defineProperty(exports, "__esModule", { value: true });
// Added by generator-teams
__export(__webpack_require__(3));
__export(__webpack_require__(4));


/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var theme_1 = __webpack_require__(0);
/**
 * Implementation of Module 1 configuration page
 */
var module1Configure = /** @class */ (function () {
    function module1Configure() {
        var _this = this;
        microsoftTeams.initialize();
        microsoftTeams.getContext(function (context) {
            theme_1.TeamsTheme.fix(context);
            var val = document.getElementById("data");
            if (context.entityId) {
                val.value = context.entityId;
            }
            _this.setValidityState(true);
        });
        microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
            var val = document.getElementById("data");
            // Calculate host dynamically to enable local debugging
            var host = "https://" + window.location.host;
            microsoftTeams.settings.setSettings({
                contentUrl: host + "/module1Tab.html?data=",
                suggestedDisplayName: 'Module 1',
                removeUrl: host + "/module1Remove.html",
                entityId: val.value
            });
            saveEvent.notifySuccess();
        });
    }
    module1Configure.prototype.setValidityState = function (val) {
        microsoftTeams.settings.setValidityState(val);
    };
    return module1Configure;
}());
exports.module1Configure = module1Configure;


/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var theme_1 = __webpack_require__(0);
/**
 * Implementation of the Module 1 content page
 */
var module1Tab = /** @class */ (function () {
    /**
     * Constructor for module1 that initializes the Microsoft Teams script and themes management
     */
    function module1Tab() {
        microsoftTeams.initialize();
        theme_1.TeamsTheme.fix();
    }
    /**
     * Method to invoke on page to start processing
     * Add your custom implementation here
     */
    module1Tab.prototype.doStuff = function () {
        microsoftTeams.getContext(function (context) {
            var element = document.getElementById('app');
            if (element) {
                element.innerHTML = "The value is: " + context.entityId;
            }
        });
    };
    return module1Tab;
}());
exports.module1Tab = module1Tab;


/***/ })
/******/ ]);
});
//# sourceMappingURL=client.js.map