/******/ (() => { // webpackBootstrap
/******/ 	"use strict";

// UNUSED EXPORTS: formatAgentError

;// ./node_modules/dompurify/dist/purify.es.mjs
/*! @license DOMPurify 3.4.13 | (c) Cure53 and other contributors | Released under the Apache license 2.0 and Mozilla Public License 2.0 | github.com/cure53/DOMPurify/blob/3.4.13/LICENSE */

function _arrayLikeToArray(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function _arrayWithHoles(r) {
  if (Array.isArray(r)) return r;
}
function _iterableToArrayLimit(r, l) {
  var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (null != t) {
    var e,
      n,
      i,
      u,
      a = [],
      f = true,
      o = false;
    try {
      if (i = (t = t.call(r)).next, 0 === l) ; else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0);
    } catch (r) {
      o = true, n = r;
    } finally {
      try {
        if (!f && null != t.return && (u = t.return(), Object(u) !== u)) return;
      } finally {
        if (o) throw n;
      }
    }
    return a;
  }
}
function _nonIterableRest() {
  throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _slicedToArray(r, e) {
  return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest();
}
function _unsupportedIterableToArray(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0;
  }
}

const entries = Object.entries,
  setPrototypeOf = Object.setPrototypeOf,
  isFrozen = Object.isFrozen,
  getPrototypeOf = Object.getPrototypeOf,
  getOwnPropertyDescriptor = Object.getOwnPropertyDescriptor;
let freeze = Object.freeze,
  seal = Object.seal,
  create = Object.create; // eslint-disable-line import/no-mutable-exports
let _ref = typeof Reflect !== 'undefined' && Reflect,
  apply = _ref.apply,
  construct = _ref.construct;
if (!freeze) {
  freeze = function freeze(x) {
    return x;
  };
}
if (!seal) {
  seal = function seal(x) {
    return x;
  };
}
if (!apply) {
  apply = function apply(func, thisArg) {
    for (var _len = arguments.length, args = new Array(_len > 2 ? _len - 2 : 0), _key = 2; _key < _len; _key++) {
      args[_key - 2] = arguments[_key];
    }
    return func.apply(thisArg, args);
  };
}
if (!construct) {
  construct = function construct(Func) {
    for (var _len2 = arguments.length, args = new Array(_len2 > 1 ? _len2 - 1 : 0), _key2 = 1; _key2 < _len2; _key2++) {
      args[_key2 - 1] = arguments[_key2];
    }
    return new Func(...args);
  };
}
const arrayForEach = unapply(Array.prototype.forEach);
const arrayLastIndexOf = unapply(Array.prototype.lastIndexOf);
const arrayPop = unapply(Array.prototype.pop);
const arrayPush = unapply(Array.prototype.push);
const arraySplice = unapply(Array.prototype.splice);
const arrayIsArray = Array.isArray;
const stringToLowerCase = unapply(String.prototype.toLowerCase);
const stringToString = unapply(String.prototype.toString);
const stringMatch = unapply(String.prototype.match);
const stringReplace = unapply(String.prototype.replace);
const stringIndexOf = unapply(String.prototype.indexOf);
const stringTrim = unapply(String.prototype.trim);
const numberToString = unapply(Number.prototype.toString);
const booleanToString = unapply(Boolean.prototype.toString);
const bigintToString = typeof BigInt === 'undefined' ? null : unapply(BigInt.prototype.toString);
const symbolToString = typeof Symbol === 'undefined' ? null : unapply(Symbol.prototype.toString);
const objectHasOwnProperty = unapply(Object.prototype.hasOwnProperty);
const objectToString = unapply(Object.prototype.toString);
const regExpTest = unapply(RegExp.prototype.test);
const typeErrorCreate = unconstruct(TypeError);
/**
 * Creates a new function that calls the given function with a specified thisArg and arguments.
 *
 * @param func - The function to be wrapped and called.
 * @returns A new function that calls the given function with a specified thisArg and arguments.
 */
function unapply(func) {
  return function (thisArg) {
    if (thisArg instanceof RegExp) {
      thisArg.lastIndex = 0;
    }
    for (var _len3 = arguments.length, args = new Array(_len3 > 1 ? _len3 - 1 : 0), _key3 = 1; _key3 < _len3; _key3++) {
      args[_key3 - 1] = arguments[_key3];
    }
    return apply(func, thisArg, args);
  };
}
/**
 * Creates a new function that constructs an instance of the given constructor function with the provided arguments.
 *
 * @param func - The constructor function to be wrapped and called.
 * @returns A new function that constructs an instance of the given constructor function with the provided arguments.
 */
function unconstruct(Func) {
  return function () {
    for (var _len4 = arguments.length, args = new Array(_len4), _key4 = 0; _key4 < _len4; _key4++) {
      args[_key4] = arguments[_key4];
    }
    return construct(Func, args);
  };
}
/**
 * Add properties to a lookup table
 *
 * @param set - The set to which elements will be added.
 * @param array - The array containing elements to be added to the set.
 * @param transformCaseFunc - An optional function to transform the case of each element before adding to the set.
 * @returns The modified set with added elements.
 */
function addToSet(set, array) {
  let transformCaseFunc = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : stringToLowerCase;
  if (setPrototypeOf) {
    // Make 'in' and truthy checks like Boolean(set.constructor)
    // independent of any properties defined on Object.prototype.
    // Prevent prototype setters from intercepting set as a this value.
    setPrototypeOf(set, null);
  }
  if (!arrayIsArray(array)) {
    return set;
  }
  let l = array.length;
  while (l--) {
    let element = array[l];
    if (typeof element === 'string') {
      const lcElement = transformCaseFunc(element);
      if (lcElement !== element) {
        // Config presets (e.g. tags.js, attrs.js) are immutable.
        if (!isFrozen(array)) {
          array[l] = lcElement;
        }
        element = lcElement;
      }
    }
    set[element] = true;
  }
  return set;
}
/**
 * Clean up an array to harden against CSPP
 *
 * @param array - The array to be cleaned.
 * @returns The cleaned version of the array
 */
function cleanArray(array) {
  for (let index = 0; index < array.length; index++) {
    const isPropertyExist = objectHasOwnProperty(array, index);
    if (!isPropertyExist) {
      array[index] = null;
    }
  }
  return array;
}
/**
 * Shallow clone an object
 *
 * @param object - The object to be cloned.
 * @returns A new object that copies the original.
 */
function clone(object) {
  const newObject = create(null);
  for (const _ref2 of entries(object)) {
    var _ref3 = _slicedToArray(_ref2, 2);
    const property = _ref3[0];
    const value = _ref3[1];
    const isPropertyExist = objectHasOwnProperty(object, property);
    if (isPropertyExist) {
      if (arrayIsArray(value)) {
        newObject[property] = cleanArray(value);
      } else if (value && typeof value === 'object' && value.constructor === Object) {
        newObject[property] = clone(value);
      } else {
        newObject[property] = value;
      }
    }
  }
  return newObject;
}
/**
 * Convert non-node values into strings without depending on direct property access.
 *
 * @param value - The value to stringify.
 * @returns A string representation of the provided value.
 */
function stringifyValue(value) {
  switch (typeof value) {
    case 'string':
      {
        return value;
      }
    case 'number':
      {
        return numberToString(value);
      }
    case 'boolean':
      {
        return booleanToString(value);
      }
    case 'bigint':
      {
        return bigintToString ? bigintToString(value) : '0';
      }
    case 'symbol':
      {
        return symbolToString ? symbolToString(value) : 'Symbol()';
      }
    case 'undefined':
      {
        return objectToString(value);
      }
    case 'function':
    case 'object':
      {
        if (value === null) {
          return objectToString(value);
        }
        const valueAsRecord = value;
        const valueToString = lookupGetter(valueAsRecord, 'toString');
        if (typeof valueToString === 'function') {
          const stringified = valueToString(valueAsRecord);
          return typeof stringified === 'string' ? stringified : objectToString(stringified);
        }
        return objectToString(value);
      }
    default:
      {
        return objectToString(value);
      }
  }
}
/**
 * This method automatically checks if the prop is function or getter and behaves accordingly.
 *
 * @param object - The object to look up the getter function in its prototype chain.
 * @param prop - The property name for which to find the getter function.
 * @returns The getter function found in the prototype chain or a fallback function.
 */
function lookupGetter(object, prop) {
  while (object !== null) {
    const desc = getOwnPropertyDescriptor(object, prop);
    if (desc) {
      if (desc.get) {
        return unapply(desc.get);
      }
      if (typeof desc.value === 'function') {
        return unapply(desc.value);
      }
    }
    object = getPrototypeOf(object);
  }
  function fallbackValue() {
    return null;
  }
  return fallbackValue;
}
function isRegex(value) {
  try {
    regExpTest(value, '');
    return true;
  } catch (_unused) {
    return false;
  }
}

const html$1 = freeze(['a', 'abbr', 'acronym', 'address', 'area', 'article', 'aside', 'audio', 'b', 'bdi', 'bdo', 'big', 'blink', 'blockquote', 'body', 'br', 'button', 'canvas', 'caption', 'center', 'cite', 'code', 'col', 'colgroup', 'content', 'data', 'datalist', 'dd', 'decorator', 'del', 'details', 'dfn', 'dialog', 'dir', 'div', 'dl', 'dt', 'element', 'em', 'fieldset', 'figcaption', 'figure', 'font', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'head', 'header', 'hgroup', 'hr', 'html', 'i', 'img', 'input', 'ins', 'kbd', 'label', 'legend', 'li', 'main', 'map', 'mark', 'marquee', 'menu', 'menuitem', 'meter', 'nav', 'nobr', 'ol', 'optgroup', 'option', 'output', 'p', 'picture', 'pre', 'progress', 'q', 'rp', 'rt', 'ruby', 's', 'samp', 'search', 'section', 'select', 'shadow', 'slot', 'small', 'source', 'spacer', 'span', 'strike', 'strong', 'style', 'sub', 'summary', 'sup', 'table', 'tbody', 'td', 'template', 'textarea', 'tfoot', 'th', 'thead', 'time', 'tr', 'track', 'tt', 'u', 'ul', 'var', 'video', 'wbr']);
const svg$1 = freeze(['svg', 'a', 'altglyph', 'altglyphdef', 'altglyphitem', 'animatecolor', 'animatemotion', 'animatetransform', 'circle', 'clippath', 'defs', 'desc', 'ellipse', 'enterkeyhint', 'exportparts', 'filter', 'font', 'g', 'glyph', 'glyphref', 'hkern', 'image', 'inputmode', 'line', 'lineargradient', 'marker', 'mask', 'metadata', 'mpath', 'part', 'path', 'pattern', 'polygon', 'polyline', 'radialgradient', 'rect', 'stop', 'style', 'switch', 'symbol', 'text', 'textpath', 'title', 'tref', 'tspan', 'view', 'vkern']);
const svgFilters = freeze(['feBlend', 'feColorMatrix', 'feComponentTransfer', 'feComposite', 'feConvolveMatrix', 'feDiffuseLighting', 'feDisplacementMap', 'feDistantLight', 'feDropShadow', 'feFlood', 'feFuncA', 'feFuncB', 'feFuncG', 'feFuncR', 'feGaussianBlur', 'feImage', 'feMerge', 'feMergeNode', 'feMorphology', 'feOffset', 'fePointLight', 'feSpecularLighting', 'feSpotLight', 'feTile', 'feTurbulence']);
// List of SVG elements that are disallowed by default.
// We still need to know them so that we can do namespace
// checks properly in case one wants to add them to
// allow-list.
const svgDisallowed = freeze(['animate', 'color-profile', 'cursor', 'discard', 'font-face', 'font-face-format', 'font-face-name', 'font-face-src', 'font-face-uri', 'foreignobject', 'hatch', 'hatchpath', 'mesh', 'meshgradient', 'meshpatch', 'meshrow', 'missing-glyph', 'script', 'set', 'solidcolor', 'unknown', 'use']);
const mathMl$1 = freeze(['math', 'menclose', 'merror', 'mfenced', 'mfrac', 'mglyph', 'mi', 'mlabeledtr', 'mmultiscripts', 'mn', 'mo', 'mover', 'mpadded', 'mphantom', 'mroot', 'mrow', 'ms', 'mspace', 'msqrt', 'mstyle', 'msub', 'msup', 'msubsup', 'mtable', 'mtd', 'mtext', 'mtr', 'munder', 'munderover', 'mprescripts']);
// Similarly to SVG, we want to know all MathML elements,
// even those that we disallow by default.
const mathMlDisallowed = freeze(['maction', 'maligngroup', 'malignmark', 'mlongdiv', 'mscarries', 'mscarry', 'msgroup', 'mstack', 'msline', 'msrow', 'semantics', 'annotation', 'annotation-xml', 'mprescripts', 'none']);
const purify_es_text = freeze(['#text']);

const html = freeze(['accept', 'action', 'align', 'alt', 'autocapitalize', 'autocomplete', 'autopictureinpicture', 'autoplay', 'background', 'bgcolor', 'border', 'capture', 'cellpadding', 'cellspacing', 'checked', 'cite', 'class', 'clear', 'color', 'cols', 'colspan', 'command', 'commandfor', 'controls', 'controlslist', 'coords', 'crossorigin', 'datetime', 'decoding', 'default', 'dir', 'disabled', 'disablepictureinpicture', 'disableremoteplayback', 'download', 'draggable', 'enctype', 'enterkeyhint', 'exportparts', 'face', 'for', 'headers', 'height', 'hidden', 'high', 'href', 'hreflang', 'id', 'inert', 'inputmode', 'integrity', 'ismap', 'kind', 'label', 'lang', 'list', 'loading', 'loop', 'low', 'max', 'maxlength', 'media', 'method', 'min', 'minlength', 'multiple', 'muted', 'name', 'nonce', 'noshade', 'novalidate', 'nowrap', 'open', 'optimum', 'part', 'pattern', 'placeholder', 'playsinline', 'popover', 'popovertarget', 'popovertargetaction', 'poster', 'preload', 'pubdate', 'radiogroup', 'readonly', 'rel', 'required', 'rev', 'reversed', 'role', 'rows', 'rowspan', 'spellcheck', 'scope', 'selected', 'shape', 'size', 'sizes', 'slot', 'span', 'srclang', 'start', 'src', 'srcset', 'step', 'style', 'summary', 'tabindex', 'title', 'translate', 'type', 'usemap', 'valign', 'value', 'width', 'wrap', 'xmlns']);
const svg = freeze(['accent-height', 'accumulate', 'additive', 'alignment-baseline', 'amplitude', 'ascent', 'attributename', 'attributetype', 'azimuth', 'basefrequency', 'baseline-shift', 'begin', 'bias', 'by', 'class', 'clip', 'clippathunits', 'clip-path', 'clip-rule', 'color', 'color-interpolation', 'color-interpolation-filters', 'color-profile', 'color-rendering', 'cx', 'cy', 'd', 'dx', 'dy', 'diffuseconstant', 'direction', 'display', 'divisor', 'dominant-baseline', 'dur', 'edgemode', 'elevation', 'end', 'exponent', 'fill', 'fill-opacity', 'fill-rule', 'filter', 'filterunits', 'flood-color', 'flood-opacity', 'font-family', 'font-size', 'font-size-adjust', 'font-stretch', 'font-style', 'font-variant', 'font-weight', 'fx', 'fy', 'g1', 'g2', 'glyph-name', 'glyphref', 'gradientunits', 'gradienttransform', 'height', 'href', 'id', 'image-rendering', 'in', 'in2', 'intercept', 'k', 'k1', 'k2', 'k3', 'k4', 'kerning', 'keypoints', 'keysplines', 'keytimes', 'lang', 'lengthadjust', 'letter-spacing', 'kernelmatrix', 'kernelunitlength', 'lighting-color', 'local', 'marker-end', 'marker-mid', 'marker-start', 'markerheight', 'markerunits', 'markerwidth', 'maskcontentunits', 'maskunits', 'max', 'mask', 'mask-type', 'media', 'method', 'mode', 'min', 'name', 'numoctaves', 'offset', 'operator', 'opacity', 'order', 'orient', 'orientation', 'origin', 'overflow', 'paint-order', 'path', 'pathlength', 'patterncontentunits', 'patterntransform', 'patternunits', 'points', 'preservealpha', 'preserveaspectratio', 'primitiveunits', 'r', 'rx', 'ry', 'radius', 'refx', 'refy', 'repeatcount', 'repeatdur', 'restart', 'result', 'rotate', 'scale', 'seed', 'shape-rendering', 'slope', 'specularconstant', 'specularexponent', 'spreadmethod', 'startoffset', 'stddeviation', 'stitchtiles', 'stop-color', 'stop-opacity', 'stroke-dasharray', 'stroke-dashoffset', 'stroke-linecap', 'stroke-linejoin', 'stroke-miterlimit', 'stroke-opacity', 'stroke', 'stroke-width', 'style', 'surfacescale', 'systemlanguage', 'tabindex', 'tablevalues', 'targetx', 'targety', 'transform', 'transform-origin', 'text-anchor', 'text-decoration', 'text-orientation', 'text-rendering', 'textlength', 'type', 'u1', 'u2', 'unicode', 'values', 'viewbox', 'visibility', 'version', 'vert-adv-y', 'vert-origin-x', 'vert-origin-y', 'width', 'word-spacing', 'wrap', 'writing-mode', 'xchannelselector', 'ychannelselector', 'x', 'x1', 'x2', 'xmlns', 'y', 'y1', 'y2', 'z', 'zoomandpan']);
const mathMl = freeze(['accent', 'accentunder', 'align', 'bevelled', 'close', 'columnalign', 'columnlines', 'columnspacing', 'columnspan', 'denomalign', 'depth', 'dir', 'display', 'displaystyle', 'encoding', 'fence', 'frame', 'height', 'href', 'id', 'largeop', 'length', 'linethickness', 'lquote', 'lspace', 'mathbackground', 'mathcolor', 'mathsize', 'mathvariant', 'maxsize', 'minsize', 'movablelimits', 'notation', 'numalign', 'open', 'rowalign', 'rowlines', 'rowspacing', 'rowspan', 'rspace', 'rquote', 'scriptlevel', 'scriptminsize', 'scriptsizemultiplier', 'selection', 'separator', 'separators', 'stretchy', 'subscriptshift', 'supscriptshift', 'symmetric', 'voffset', 'width', 'xmlns']);
const xml = freeze(['xlink:href', 'xml:id', 'xlink:title', 'xml:space', 'xmlns:xlink']);

const MUSTACHE_EXPR = seal(/{{[\w\W]*|^[\w\W]*}}/g);
const ERB_EXPR = seal(/<%[\w\W]*|^[\w\W]*%>/g);
const TMPLIT_EXPR = seal(/\${[\w\W]*/g);
const DATA_ATTR = seal(/^data-[\-\w.\u00B7-\uFFFF]+$/); // eslint-disable-line no-useless-escape
const ARIA_ATTR = seal(/^aria-[\-\w]+$/); // eslint-disable-line no-useless-escape
const IS_ALLOWED_URI = seal(/^(?:(?:(?:f|ht)tps?|mailto|tel|callto|sms|cid|xmpp|matrix):|[^a-z]|[a-z+.\-]+(?:[^a-z+.\-:]|$))/i // eslint-disable-line no-useless-escape
);
const IS_SCRIPT_OR_DATA = seal(/^(?:\w+script|data):/i);
const ATTR_WHITESPACE = seal(/[\u0000-\u0020\u00A0\u1680\u180E\u2000-\u2029\u205F\u3000]/g // eslint-disable-line no-control-regex
);
const DOCTYPE_NAME = seal(/^html$/i);
const CUSTOM_ELEMENT = seal(/^[a-z][.\w]*(-[.\w]+)+$/i);
// Markup-significant character probes used by _sanitizeElements.
// Shared module-level instances are safe despite the sticky /g flags:
// unapply() resets lastIndex for RegExp receivers before every call.
const ELEMENT_MARKUP_PROBE = seal(/<[/\w!]/g);
const COMMENT_MARKUP_PROBE = seal(/<[/\w]/g);
const FALLBACK_TAG_CLOSE = seal(/<\/no(script|embed|frames)/i);
const SELF_CLOSING_TAG = seal(/\/>/i);

// https://developer.mozilla.org/en-US/docs/Web/API/Node/nodeType
const NODE_TYPE = {
  element: 1,
  attribute: 2,
  text: 3,
  cdataSection: 4,
  entityReference: 5,
  // Deprecated
  entityNode: 6,
  // Deprecated
  processingInstruction: 7,
  comment: 8,
  document: 9,
  documentType: 10,
  documentFragment: 11,
  notation: 12 // Deprecated
};
const getGlobal = function getGlobal() {
  return typeof window === 'undefined' ? null : window;
};
/**
 * Creates a no-op policy for internal use only.
 * Don't export this function outside this module!
 * @param trustedTypes The policy factory.
 * @param purifyHostElement The Script element used to load DOMPurify (to determine policy name suffix).
 * @return The policy created (or null, if Trusted Types
 * are not supported or creating the policy failed).
 */
const _createTrustedTypesPolicy = function _createTrustedTypesPolicy(trustedTypes, purifyHostElement) {
  if (typeof trustedTypes !== 'object' || typeof trustedTypes.createPolicy !== 'function') {
    return null;
  }
  // Allow the callers to control the unique policy name
  // by adding a data-tt-policy-suffix to the script element with the DOMPurify.
  // Policy creation with duplicate names throws in Trusted Types.
  let suffix = null;
  const ATTR_NAME = 'data-tt-policy-suffix';
  if (purifyHostElement && purifyHostElement.hasAttribute(ATTR_NAME)) {
    suffix = purifyHostElement.getAttribute(ATTR_NAME);
  }
  const policyName = 'dompurify' + (suffix ? '#' + suffix : '');
  try {
    return trustedTypes.createPolicy(policyName, {
      createHTML(html) {
        return html;
      },
      createScriptURL(scriptUrl) {
        return scriptUrl;
      }
    });
  } catch (_) {
    // Policy creation failed (most likely another DOMPurify script has
    // already run). Skip creating the policy, as this will only cause errors
    // if TT are enforced.
    console.warn('TrustedTypes policy ' + policyName + ' could not be created.');
    return null;
  }
};
const _createHooksMap = function _createHooksMap() {
  return {
    afterSanitizeAttributes: [],
    afterSanitizeElements: [],
    afterSanitizeShadowDOM: [],
    beforeSanitizeAttributes: [],
    beforeSanitizeElements: [],
    beforeSanitizeShadowDOM: [],
    uponSanitizeAttribute: [],
    uponSanitizeElement: [],
    uponSanitizeShadowNode: []
  };
};
/**
 * Resolve a set-valued configuration option: a fresh set built from
 * cfg[key] when it is an own array property (seeded with a clone of
 * options.base when given, case-normalized via options.transform),
 * the fallback set otherwise.
 *
 * @param cfg the cloned, prototype-free configuration object
 * @param key the configuration property to read
 * @param fallback the set to use when the option is absent or not an array
 * @param options transform and optional base set to merge into
 * @returns the resolved set
 */
const _resolveSetOption = function _resolveSetOption(cfg, key, fallback, options) {
  return objectHasOwnProperty(cfg, key) && arrayIsArray(cfg[key]) ? addToSet(options.base ? clone(options.base) : {}, cfg[key], options.transform) : fallback;
};
function createDOMPurify() {
  let window = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : getGlobal();
  const DOMPurify = root => createDOMPurify(root);
  DOMPurify.version = '3.4.13';
  DOMPurify.removed = [];
  if (!window || !window.document || window.document.nodeType !== NODE_TYPE.document || !window.Element) {
    // Not running in a browser, provide a factory function
    // so that you can pass your own Window
    DOMPurify.isSupported = false;
    return DOMPurify;
  }
  let document = window.document;
  const originalDocument = document;
  const currentScript = originalDocument.currentScript;
  window.DocumentFragment;
    const HTMLTemplateElement = window.HTMLTemplateElement,
    Node = window.Node,
    Element = window.Element,
    NodeFilter = window.NodeFilter,
    _window$NamedNodeMap = window.NamedNodeMap;
    _window$NamedNodeMap === void 0 ? window.NamedNodeMap || window.MozNamedAttrMap : _window$NamedNodeMap;
    window.HTMLFormElement;
    const DOMParser = window.DOMParser,
    trustedTypes = window.trustedTypes;
  const ElementPrototype = Element.prototype;
  const cloneNode = lookupGetter(ElementPrototype, 'cloneNode');
  const remove = lookupGetter(ElementPrototype, 'remove');
  const getNextSibling = lookupGetter(ElementPrototype, 'nextSibling');
  const getChildNodes = lookupGetter(ElementPrototype, 'childNodes');
  const getParentNode = lookupGetter(ElementPrototype, 'parentNode');
  const getShadowRoot = lookupGetter(ElementPrototype, 'shadowRoot');
  const getAttributes = lookupGetter(ElementPrototype, 'attributes');
  const getNodeType = Node && Node.prototype ? lookupGetter(Node.prototype, 'nodeType') : null;
  const getNodeName = Node && Node.prototype ? lookupGetter(Node.prototype, 'nodeName') : null;
  const getOwnerDocument = Node && Node.prototype ? lookupGetter(Node.prototype, 'ownerDocument') : null;
  // As per issue #47, the web-components registry is inherited by a
  // new document created via createHTMLDocument. As per the spec
  // (http://w3c.github.io/webcomponents/spec/custom/#creating-and-passing-registries)
  // a new empty registry is used when creating a template contents owner
  // document, so we use that as our parent document to ensure nothing
  // is inherited.
  if (typeof HTMLTemplateElement === 'function') {
    const template = document.createElement('template');
    if (template.content && template.content.ownerDocument) {
      document = template.content.ownerDocument;
    }
  }
  let trustedTypesPolicy;
  let emptyHTML = '';
  // The instance's own internal Trusted Types policy. Unlike a caller-supplied
  // `TRUSTED_TYPES_POLICY`, this is created at most once — Trusted Types throws
  // on duplicate policy names — and is the only policy allowed to persist
  // across configurations and survive `clearConfig()`.
  let defaultTrustedTypesPolicy;
  let defaultTrustedTypesPolicyResolved = false;
  // Tracks whether we are already inside a call to the configured Trusted Types
  // policy (`createHTML` or `createScriptURL`). If a supplied policy callback
  // itself calls `DOMPurify.sanitize` (the cause of #1422), `sanitize` would
  // re-enter the policy and recurse until the stack overflows. We detect that
  // re-entry and throw a clear, actionable error instead. The guard is shared
  // across both callbacks, because either one re-entering `sanitize` triggers
  // the same unbounded recursion.
  let IN_TRUSTED_TYPES_POLICY = 0;
  const _assertNotInTrustedTypesPolicy = function _assertNotInTrustedTypesPolicy() {
    if (IN_TRUSTED_TYPES_POLICY > 0) {
      throw typeErrorCreate('A configured TRUSTED_TYPES_POLICY callback (createHTML or ' + 'createScriptURL) must not call DOMPurify.sanitize, as that causes ' + 'infinite recursion. Do not pass a policy whose callbacks wrap ' + 'DOMPurify as TRUSTED_TYPES_POLICY; see the "DOMPurify and Trusted ' + 'Types" section of the README.');
    }
  };
  const _createTrustedHTML = function _createTrustedHTML(html) {
    _assertNotInTrustedTypesPolicy();
    IN_TRUSTED_TYPES_POLICY++;
    try {
      return trustedTypesPolicy.createHTML(html);
    } finally {
      IN_TRUSTED_TYPES_POLICY--;
    }
  };
  const _createTrustedScriptURL = function _createTrustedScriptURL(scriptUrl) {
    _assertNotInTrustedTypesPolicy();
    IN_TRUSTED_TYPES_POLICY++;
    try {
      return trustedTypesPolicy.createScriptURL(scriptUrl);
    } finally {
      IN_TRUSTED_TYPES_POLICY--;
    }
  };
  // Lazily resolve (and cache) the instance's internal default policy.
  // Resolution is attempted at most once: a successful `createPolicy` cannot be
  // repeated (Trusted Types throws on duplicate names), and a failed or
  // unsupported attempt must not be retried on every parse.
  const _getDefaultTrustedTypesPolicy = function _getDefaultTrustedTypesPolicy() {
    if (!defaultTrustedTypesPolicyResolved) {
      defaultTrustedTypesPolicy = _createTrustedTypesPolicy(trustedTypes, currentScript);
      defaultTrustedTypesPolicyResolved = true;
    }
    return defaultTrustedTypesPolicy;
  };
  const _document = document,
    implementation = _document.implementation,
    createNodeIterator = _document.createNodeIterator,
    createDocumentFragment = _document.createDocumentFragment,
    getElementsByTagName = _document.getElementsByTagName;
  const importNode = originalDocument.importNode;
  let hooks = _createHooksMap();
  /**
   * Expose whether this browser supports running the full DOMPurify.
   */
  DOMPurify.isSupported = typeof entries === 'function' && typeof getParentNode === 'function' && implementation && implementation.createHTMLDocument !== undefined;
  const MUSTACHE_EXPR$1 = MUSTACHE_EXPR,
    ERB_EXPR$1 = ERB_EXPR,
    TMPLIT_EXPR$1 = TMPLIT_EXPR,
    DATA_ATTR$1 = DATA_ATTR,
    ARIA_ATTR$1 = ARIA_ATTR,
    IS_SCRIPT_OR_DATA$1 = IS_SCRIPT_OR_DATA,
    ATTR_WHITESPACE$1 = ATTR_WHITESPACE,
    CUSTOM_ELEMENT$1 = CUSTOM_ELEMENT;
  let IS_ALLOWED_URI$1 = IS_ALLOWED_URI;
  /**
   * We consider the elements and attributes below to be safe. Ideally
   * don't add any new ones but feel free to remove unwanted ones.
   */
  /* allowed element names */
  let ALLOWED_TAGS = null;
  const DEFAULT_ALLOWED_TAGS = addToSet({}, [...html$1, ...svg$1, ...svgFilters, ...mathMl$1, ...purify_es_text]);
  /* Allowed attribute names */
  let ALLOWED_ATTR = null;
  const DEFAULT_ALLOWED_ATTR = addToSet({}, [...html, ...svg, ...mathMl, ...xml]);
  /*
   * Configure how DOMPurify should handle custom elements and their attributes as well as customized built-in elements.
   * @property {RegExp|Function|null} tagNameCheck one of [null, regexPattern, predicate]. Default: `null` (disallow any custom elements)
   * @property {RegExp|Function|null} attributeNameCheck one of [null, regexPattern, predicate]. Default: `null` (disallow any attributes not on the allow list)
   * @property {boolean} allowCustomizedBuiltInElements allow custom elements derived from built-ins if they pass CUSTOM_ELEMENT_HANDLING.tagNameCheck. Default: `false`.
   */
  let CUSTOM_ELEMENT_HANDLING = Object.seal(create(null, {
    tagNameCheck: {
      writable: true,
      configurable: false,
      enumerable: true,
      value: null
    },
    attributeNameCheck: {
      writable: true,
      configurable: false,
      enumerable: true,
      value: null
    },
    allowCustomizedBuiltInElements: {
      writable: true,
      configurable: false,
      enumerable: true,
      value: false
    }
  }));
  /* Explicitly forbidden tags (overrides ALLOWED_TAGS/ADD_TAGS) */
  let FORBID_TAGS = null;
  /* Explicitly forbidden attributes (overrides ALLOWED_ATTR/ADD_ATTR) */
  let FORBID_ATTR = null;
  /* Config object to store ADD_TAGS/ADD_ATTR functions (when used as functions) */
  const EXTRA_ELEMENT_HANDLING = Object.seal(create(null, {
    tagCheck: {
      writable: true,
      configurable: false,
      enumerable: true,
      value: null
    },
    attributeCheck: {
      writable: true,
      configurable: false,
      enumerable: true,
      value: null
    }
  }));
  /* Decide if ARIA attributes are okay */
  let ALLOW_ARIA_ATTR = true;
  /* Decide if custom data attributes are okay */
  let ALLOW_DATA_ATTR = true;
  /* Decide if unknown protocols are okay */
  let ALLOW_UNKNOWN_PROTOCOLS = false;
  /* Decide if self-closing tags in attributes are allowed.
   * Usually removed due to a mXSS issue in jQuery 3.0 */
  let ALLOW_SELF_CLOSE_IN_ATTR = true;
  /* Output should be safe for common template engines.
   * This means, DOMPurify removes data attributes, mustaches and ERB
   */
  let SAFE_FOR_TEMPLATES = false;
  /* Output should be safe even for XML used within HTML and alike.
   * This means, DOMPurify removes comments when containing risky content.
   */
  let SAFE_FOR_XML = true;
  /* Decide if document with <html>... should be returned */
  let WHOLE_DOCUMENT = false;
  /* Track whether config is already set on this instance of DOMPurify. */
  let SET_CONFIG = false;
  /* Pristine allowlist bindings captured at setConfig() time. On the
   * persistent-config path sanitize() restores the sets from these before
   * the per-walk hook clone-guard, so a hook's in-call widening cannot
   * carry across calls. Null until setConfig() is called; reset by
   * clearConfig(). */
  let SET_CONFIG_ALLOWED_TAGS = null;
  let SET_CONFIG_ALLOWED_ATTR = null;
  /* Decide if all elements (e.g. style, script) must be children of
   * document.body. By default, browsers might move them to document.head */
  let FORCE_BODY = false;
  /* Decide if a DOM `HTMLBodyElement` should be returned, instead of a html
   * string (or a TrustedHTML object if Trusted Types are supported).
   * If `WHOLE_DOCUMENT` is enabled a `HTMLHtmlElement` will be returned instead
   */
  let RETURN_DOM = false;
  /* Decide if a DOM `DocumentFragment` should be returned, instead of a html
   * string  (or a TrustedHTML object if Trusted Types are supported) */
  let RETURN_DOM_FRAGMENT = false;
  /* Try to return a Trusted Type object instead of a string, return a string in
   * case Trusted Types are not supported  */
  let RETURN_TRUSTED_TYPE = false;
  /* Output should be free from DOM clobbering attacks?
   * This sanitizes markups named with colliding, clobberable built-in DOM APIs.
   */
  let SANITIZE_DOM = true;
  /* Achieve full DOM Clobbering protection by isolating the namespace of named
   * properties and JS variables, mitigating attacks that abuse the HTML/DOM spec rules.
   *
   * HTML/DOM spec rules that enable DOM Clobbering:
   *   - Named Access on Window (§7.3.3)
   *   - DOM Tree Accessors (§3.1.5)
   *   - Form Element Parent-Child Relations (§4.10.3)
   *   - Iframe srcdoc / Nested WindowProxies (§4.8.5)
   *   - HTMLCollection (§4.2.10.2)
   *
   * Namespace isolation is implemented by prefixing `id` and `name` attributes
   * with a constant string, i.e., `user-content-`
   */
  let SANITIZE_NAMED_PROPS = false;
  const SANITIZE_NAMED_PROPS_PREFIX = 'user-content-';
  /* Keep element content when removing element? */
  let KEEP_CONTENT = true;
  /* If a `Node` is passed to sanitize(), then performs sanitization in-place instead
   * of importing it into a new Document and returning a sanitized copy */
  let IN_PLACE = false;
  /* Allow usage of profiles like html, svg and mathMl */
  let USE_PROFILES = {};
  /* Tags to ignore content of when KEEP_CONTENT is true */
  let FORBID_CONTENTS = null;
  const DEFAULT_FORBID_CONTENTS = addToSet({}, ['annotation-xml', 'audio', 'colgroup', 'desc', 'foreignobject', 'head', 'iframe', 'math', 'mi', 'mn', 'mo', 'ms', 'mtext', 'noembed', 'noframes', 'noscript', 'plaintext', 'script',
  // <selectedcontent> mirrors the selected <option>'s subtree, cloned by
  // the UA (customizable <select>) — including any on* handlers — and the
  // engine re-mirrors synchronously whenever a removal changes which
  // option/selectedcontent is current, even inside DOMPurify's inert
  // DOMParser document. Hoisting its children on removal re-inserts a fresh
  // mirror target ahead of the walk, which the engine refills, looping
  // forever (DoS) and amplifying output. Dropping its content on removal
  // (rather than hoisting) breaks that cascade; the content is a duplicate
  // of the option, which is sanitized on its own. See campaign-3 F1/F6.
  'selectedcontent', 'style', 'svg', 'template', 'thead', 'title', 'video', 'xmp']);
  /* Tags that are safe for data: URIs */
  let DATA_URI_TAGS = null;
  const DEFAULT_DATA_URI_TAGS = addToSet({}, ['audio', 'video', 'img', 'source', 'image', 'track']);
  /* Attributes safe for values like "javascript:" */
  let URI_SAFE_ATTRIBUTES = null;
  const DEFAULT_URI_SAFE_ATTRIBUTES = addToSet({}, ['alt', 'class', 'for', 'id', 'label', 'name', 'pattern', 'placeholder', 'role', 'summary', 'title', 'value', 'style', 'xmlns']);
  const MATHML_NAMESPACE = 'http://www.w3.org/1998/Math/MathML';
  const SVG_NAMESPACE = 'http://www.w3.org/2000/svg';
  const HTML_NAMESPACE = 'http://www.w3.org/1999/xhtml';
  /* Document namespace */
  let NAMESPACE = HTML_NAMESPACE;
  let IS_EMPTY_INPUT = false;
  /* Allowed XHTML+XML namespaces */
  let ALLOWED_NAMESPACES = null;
  const DEFAULT_ALLOWED_NAMESPACES = addToSet({}, [MATHML_NAMESPACE, SVG_NAMESPACE, HTML_NAMESPACE], stringToString);
  const DEFAULT_MATHML_TEXT_INTEGRATION_POINTS = freeze(['mi', 'mo', 'mn', 'ms', 'mtext']);
  let MATHML_TEXT_INTEGRATION_POINTS = addToSet({}, DEFAULT_MATHML_TEXT_INTEGRATION_POINTS);
  const DEFAULT_HTML_INTEGRATION_POINTS = freeze(['annotation-xml']);
  let HTML_INTEGRATION_POINTS = addToSet({}, DEFAULT_HTML_INTEGRATION_POINTS);
  // Certain elements are allowed in both SVG and HTML
  // namespace. We need to specify them explicitly
  // so that they don't get erroneously deleted from
  // HTML namespace.
  const COMMON_SVG_AND_HTML_ELEMENTS = addToSet({}, ['title', 'style', 'font', 'a', 'script']);
  /* Parsing of strict XHTML documents */
  let PARSER_MEDIA_TYPE = null;
  const SUPPORTED_PARSER_MEDIA_TYPES = ['application/xhtml+xml', 'text/html'];
  const DEFAULT_PARSER_MEDIA_TYPE = 'text/html';
  let transformCaseFunc = null;
  /* Keep a reference to config to pass to hooks */
  let CONFIG = null;
  /* Ideally, do not touch anything below this line */
  /* ______________________________________________ */
  const formElement = document.createElement('form');
  const isRegexOrFunction = function isRegexOrFunction(testValue) {
    return testValue instanceof RegExp || testValue instanceof Function;
  };
  /**
   * _parseConfig
   *
   * @param cfg optional config literal
   */
  // eslint-disable-next-line complexity
  const _parseConfig = function _parseConfig() {
    let cfg = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
    if (CONFIG && CONFIG === cfg) {
      return;
    }
    /* Shield configuration object from tampering */
    if (!cfg || typeof cfg !== 'object') {
      cfg = {};
    }
    /* Shield configuration object from prototype pollution */
    cfg = clone(cfg);
    PARSER_MEDIA_TYPE =
    // eslint-disable-next-line unicorn/prefer-includes
    SUPPORTED_PARSER_MEDIA_TYPES.indexOf(cfg.PARSER_MEDIA_TYPE) === -1 ? DEFAULT_PARSER_MEDIA_TYPE : cfg.PARSER_MEDIA_TYPE;
    // HTML tags and attributes are not case-sensitive, converting to lowercase. Keeping XHTML as is.
    transformCaseFunc = PARSER_MEDIA_TYPE === 'application/xhtml+xml' ? stringToString : stringToLowerCase;
    /* Set configuration parameters */
    ALLOWED_TAGS = _resolveSetOption(cfg, 'ALLOWED_TAGS', DEFAULT_ALLOWED_TAGS, {
      transform: transformCaseFunc
    });
    ALLOWED_ATTR = _resolveSetOption(cfg, 'ALLOWED_ATTR', DEFAULT_ALLOWED_ATTR, {
      transform: transformCaseFunc
    });
    ALLOWED_NAMESPACES = _resolveSetOption(cfg, 'ALLOWED_NAMESPACES', DEFAULT_ALLOWED_NAMESPACES, {
      transform: stringToString
    });
    URI_SAFE_ATTRIBUTES = _resolveSetOption(cfg, 'ADD_URI_SAFE_ATTR', DEFAULT_URI_SAFE_ATTRIBUTES, {
      transform: transformCaseFunc,
      base: DEFAULT_URI_SAFE_ATTRIBUTES
    });
    DATA_URI_TAGS = _resolveSetOption(cfg, 'ADD_DATA_URI_TAGS', DEFAULT_DATA_URI_TAGS, {
      transform: transformCaseFunc,
      base: DEFAULT_DATA_URI_TAGS
    });
    FORBID_CONTENTS = _resolveSetOption(cfg, 'FORBID_CONTENTS', DEFAULT_FORBID_CONTENTS, {
      transform: transformCaseFunc
    });
    FORBID_TAGS = _resolveSetOption(cfg, 'FORBID_TAGS', clone({}), {
      transform: transformCaseFunc
    });
    FORBID_ATTR = _resolveSetOption(cfg, 'FORBID_ATTR', clone({}), {
      transform: transformCaseFunc
    });
    USE_PROFILES = objectHasOwnProperty(cfg, 'USE_PROFILES') ? cfg.USE_PROFILES && typeof cfg.USE_PROFILES === 'object' ? clone(cfg.USE_PROFILES) : cfg.USE_PROFILES : false;
    ALLOW_ARIA_ATTR = cfg.ALLOW_ARIA_ATTR !== false; // Default true
    ALLOW_DATA_ATTR = cfg.ALLOW_DATA_ATTR !== false; // Default true
    ALLOW_UNKNOWN_PROTOCOLS = cfg.ALLOW_UNKNOWN_PROTOCOLS || false; // Default false
    ALLOW_SELF_CLOSE_IN_ATTR = cfg.ALLOW_SELF_CLOSE_IN_ATTR !== false; // Default true
    SAFE_FOR_TEMPLATES = cfg.SAFE_FOR_TEMPLATES || false; // Default false
    SAFE_FOR_XML = cfg.SAFE_FOR_XML !== false; // Default true
    WHOLE_DOCUMENT = cfg.WHOLE_DOCUMENT || false; // Default false
    RETURN_DOM = cfg.RETURN_DOM || false; // Default false
    RETURN_DOM_FRAGMENT = cfg.RETURN_DOM_FRAGMENT || false; // Default false
    RETURN_TRUSTED_TYPE = cfg.RETURN_TRUSTED_TYPE || false; // Default false
    FORCE_BODY = cfg.FORCE_BODY || false; // Default false
    SANITIZE_DOM = cfg.SANITIZE_DOM !== false; // Default true
    SANITIZE_NAMED_PROPS = cfg.SANITIZE_NAMED_PROPS || false; // Default false
    KEEP_CONTENT = cfg.KEEP_CONTENT !== false; // Default true
    IN_PLACE = cfg.IN_PLACE || false; // Default false
    IS_ALLOWED_URI$1 = isRegex(cfg.ALLOWED_URI_REGEXP) ? cfg.ALLOWED_URI_REGEXP : IS_ALLOWED_URI; // Default regexp
    NAMESPACE = typeof cfg.NAMESPACE === 'string' ? cfg.NAMESPACE : HTML_NAMESPACE; // Default HTML namespace
    MATHML_TEXT_INTEGRATION_POINTS = objectHasOwnProperty(cfg, 'MATHML_TEXT_INTEGRATION_POINTS') && cfg.MATHML_TEXT_INTEGRATION_POINTS && typeof cfg.MATHML_TEXT_INTEGRATION_POINTS === 'object' ? clone(cfg.MATHML_TEXT_INTEGRATION_POINTS) : addToSet({}, DEFAULT_MATHML_TEXT_INTEGRATION_POINTS); // Default built-in map
    HTML_INTEGRATION_POINTS = objectHasOwnProperty(cfg, 'HTML_INTEGRATION_POINTS') && cfg.HTML_INTEGRATION_POINTS && typeof cfg.HTML_INTEGRATION_POINTS === 'object' ? clone(cfg.HTML_INTEGRATION_POINTS) : addToSet({}, DEFAULT_HTML_INTEGRATION_POINTS); // Default built-in map
    const customElementHandling = objectHasOwnProperty(cfg, 'CUSTOM_ELEMENT_HANDLING') && cfg.CUSTOM_ELEMENT_HANDLING && typeof cfg.CUSTOM_ELEMENT_HANDLING === 'object' ? clone(cfg.CUSTOM_ELEMENT_HANDLING) : create(null);
    CUSTOM_ELEMENT_HANDLING = create(null);
    if (objectHasOwnProperty(customElementHandling, 'tagNameCheck') && isRegexOrFunction(customElementHandling.tagNameCheck)) {
      CUSTOM_ELEMENT_HANDLING.tagNameCheck = customElementHandling.tagNameCheck; // Default undefined
    }
    if (objectHasOwnProperty(customElementHandling, 'attributeNameCheck') && isRegexOrFunction(customElementHandling.attributeNameCheck)) {
      CUSTOM_ELEMENT_HANDLING.attributeNameCheck = customElementHandling.attributeNameCheck; // Default undefined
    }
    if (objectHasOwnProperty(customElementHandling, 'allowCustomizedBuiltInElements') && typeof customElementHandling.allowCustomizedBuiltInElements === 'boolean') {
      CUSTOM_ELEMENT_HANDLING.allowCustomizedBuiltInElements = customElementHandling.allowCustomizedBuiltInElements; // Default undefined
    }
    seal(CUSTOM_ELEMENT_HANDLING);
    if (SAFE_FOR_TEMPLATES) {
      ALLOW_DATA_ATTR = false;
    }
    if (RETURN_DOM_FRAGMENT) {
      RETURN_DOM = true;
    }
    /* Parse profile info */
    if (USE_PROFILES) {
      ALLOWED_TAGS = addToSet({}, purify_es_text);
      ALLOWED_ATTR = create(null);
      if (USE_PROFILES.html === true) {
        addToSet(ALLOWED_TAGS, html$1);
        addToSet(ALLOWED_ATTR, html);
      }
      if (USE_PROFILES.svg === true) {
        addToSet(ALLOWED_TAGS, svg$1);
        addToSet(ALLOWED_ATTR, svg);
        addToSet(ALLOWED_ATTR, xml);
      }
      if (USE_PROFILES.svgFilters === true) {
        addToSet(ALLOWED_TAGS, svgFilters);
        addToSet(ALLOWED_ATTR, svg);
        addToSet(ALLOWED_ATTR, xml);
      }
      if (USE_PROFILES.mathMl === true) {
        addToSet(ALLOWED_TAGS, mathMl$1);
        addToSet(ALLOWED_ATTR, mathMl);
        addToSet(ALLOWED_ATTR, xml);
      }
    }
    /* Always reset function-based ADD_TAGS / ADD_ATTR checks to prevent
     * leaking across calls when switching from function to array config */
    EXTRA_ELEMENT_HANDLING.tagCheck = null;
    EXTRA_ELEMENT_HANDLING.attributeCheck = null;
    /* Merge configuration parameters */
    if (objectHasOwnProperty(cfg, 'ADD_TAGS')) {
      if (typeof cfg.ADD_TAGS === 'function') {
        EXTRA_ELEMENT_HANDLING.tagCheck = cfg.ADD_TAGS;
      } else if (arrayIsArray(cfg.ADD_TAGS)) {
        if (ALLOWED_TAGS === DEFAULT_ALLOWED_TAGS) {
          ALLOWED_TAGS = clone(ALLOWED_TAGS);
        }
        addToSet(ALLOWED_TAGS, cfg.ADD_TAGS, transformCaseFunc);
      }
    }
    if (objectHasOwnProperty(cfg, 'ADD_ATTR')) {
      if (typeof cfg.ADD_ATTR === 'function') {
        EXTRA_ELEMENT_HANDLING.attributeCheck = cfg.ADD_ATTR;
      } else if (arrayIsArray(cfg.ADD_ATTR)) {
        if (ALLOWED_ATTR === DEFAULT_ALLOWED_ATTR) {
          ALLOWED_ATTR = clone(ALLOWED_ATTR);
        }
        addToSet(ALLOWED_ATTR, cfg.ADD_ATTR, transformCaseFunc);
      }
    }
    if (objectHasOwnProperty(cfg, 'ADD_URI_SAFE_ATTR') && arrayIsArray(cfg.ADD_URI_SAFE_ATTR)) {
      addToSet(URI_SAFE_ATTRIBUTES, cfg.ADD_URI_SAFE_ATTR, transformCaseFunc);
    }
    if (objectHasOwnProperty(cfg, 'FORBID_CONTENTS') && arrayIsArray(cfg.FORBID_CONTENTS)) {
      if (FORBID_CONTENTS === DEFAULT_FORBID_CONTENTS) {
        FORBID_CONTENTS = clone(FORBID_CONTENTS);
      }
      addToSet(FORBID_CONTENTS, cfg.FORBID_CONTENTS, transformCaseFunc);
    }
    if (objectHasOwnProperty(cfg, 'ADD_FORBID_CONTENTS') && arrayIsArray(cfg.ADD_FORBID_CONTENTS)) {
      if (FORBID_CONTENTS === DEFAULT_FORBID_CONTENTS) {
        FORBID_CONTENTS = clone(FORBID_CONTENTS);
      }
      addToSet(FORBID_CONTENTS, cfg.ADD_FORBID_CONTENTS, transformCaseFunc);
    }
    /* Add #text in case KEEP_CONTENT is set to true */
    if (KEEP_CONTENT) {
      ALLOWED_TAGS['#text'] = true;
    }
    /* Add html, head and body to ALLOWED_TAGS in case WHOLE_DOCUMENT is true */
    if (WHOLE_DOCUMENT) {
      addToSet(ALLOWED_TAGS, ['html', 'head', 'body']);
    }
    /* Add tbody to ALLOWED_TAGS in case tables are permitted, see #286, #365 */
    if (ALLOWED_TAGS.table) {
      addToSet(ALLOWED_TAGS, ['tbody']);
      delete FORBID_TAGS.tbody;
    }
    // Re-derive the active Trusted Types policy from this configuration on
    // every parse. The active policy must never be sticky closure state that
    // outlives the config that set it: a caller-supplied policy left in place
    // after `clearConfig()` — or after a later call that supplied none, or
    // `TRUSTED_TYPES_POLICY: null` — could sign a subsequent "default"
    // `RETURN_TRUSTED_TYPE` result with a foreign, possibly unsafe policy.
    // See GHSA-vxr8-fq34-vvx9.
    if (cfg.TRUSTED_TYPES_POLICY) {
      if (typeof cfg.TRUSTED_TYPES_POLICY.createHTML !== 'function') {
        throw typeErrorCreate('TRUSTED_TYPES_POLICY configuration option must provide a "createHTML" hook.');
      }
      if (typeof cfg.TRUSTED_TYPES_POLICY.createScriptURL !== 'function') {
        throw typeErrorCreate('TRUSTED_TYPES_POLICY configuration option must provide a "createScriptURL" hook.');
      }
      // A caller-supplied policy applies to this configuration only.
      const previousTrustedTypesPolicy = trustedTypesPolicy;
      trustedTypesPolicy = cfg.TRUSTED_TYPES_POLICY;
      // Sign local variables required by `sanitize`. If the supplied policy's
      // `createHTML` is circular (i.e. it calls `DOMPurify.sanitize`), this
      // throws via the re-entrancy guard. Restore the previous policy first so
      // the instance is not left in a poisoned state. See #1422.
      try {
        emptyHTML = _createTrustedHTML('');
      } catch (error) {
        trustedTypesPolicy = previousTrustedTypesPolicy;
        throw error;
      }
    } else if (cfg.TRUSTED_TYPES_POLICY === null) {
      // Explicit opt-out for this call: perform no Trusted Types signing and
      // create nothing (so a strict `trusted-types` CSP that disallows a
      // `dompurify` policy can still call `sanitize` from inside its own
      // policy — see #1422). Resetting to `undefined` rather than a sticky
      // `null` also drops any previously retained caller policy, so it cannot
      // resurface on a later call, while still allowing the next config-less
      // call to restore the internal default policy. See GHSA-vxr8-fq34-vvx9.
      trustedTypesPolicy = undefined;
      emptyHTML = '';
    } else {
      // No policy supplied: keep the currently active policy if one is set — a
      // previously supplied policy is intentionally sticky across config-less
      // calls — otherwise fall back to the instance's own internal policy,
      // created at most once. (A policy supplied for a *single* call still
      // lingers by design; what must not linger is a policy whose configuration
      // has been torn down via `clearConfig()`, which restores the default.)
      if (trustedTypesPolicy === undefined) {
        trustedTypesPolicy = _getDefaultTrustedTypesPolicy();
      }
      // Sign internal variables only when a policy is active. A falsy policy
      // (Trusted Types unsupported, creation failed, or an explicit opt-out)
      // leaves `emptyHTML` as a plain string, so we never call `.createHTML` on
      // a non-policy and throw. See #1422.
      if (trustedTypesPolicy && typeof emptyHTML === 'string') {
        emptyHTML = _createTrustedHTML('');
      }
    }
    // Prevent further manipulation of configuration.
    // Not available in IE8, Safari 5, etc.
    if (freeze) {
      freeze(cfg);
    }
    CONFIG = cfg;
  };
  /* Keep track of all possible SVG and MathML tags
   * so that we can perform the namespace checks
   * correctly. */
  const ALL_SVG_TAGS = addToSet({}, [...svg$1, ...svgFilters, ...svgDisallowed]);
  const ALL_MATHML_TAGS = addToSet({}, [...mathMl$1, ...mathMlDisallowed]);
  /**
   * Namespace rules for an element in the SVG namespace.
   *
   * @param tagName the element's lowercase tag name
   * @param parent the (possibly simulated) parent node
   * @param parentTagName the parent's lowercase tag name
   * @returns true if a spec-compliant parser could produce this element
   */
  const _checkSvgNamespace = function _checkSvgNamespace(tagName, parent, parentTagName) {
    // The only way to switch from HTML namespace to SVG
    // is via <svg>. If it happens via any other tag, then
    // it should be killed.
    if (parent.namespaceURI === HTML_NAMESPACE) {
      return tagName === 'svg';
    }
    // The only way to switch from MathML to SVG is via <svg>
    // if the parent is either <annotation-xml> or a MathML
    // text integration point.
    if (parent.namespaceURI === MATHML_NAMESPACE) {
      return tagName === 'svg' && (parentTagName === 'annotation-xml' || MATHML_TEXT_INTEGRATION_POINTS[parentTagName]);
    }
    // We only allow elements that are defined in SVG
    // spec. All others are disallowed in SVG namespace.
    return Boolean(ALL_SVG_TAGS[tagName]);
  };
  /**
   * Namespace rules for an element in the MathML namespace.
   *
   * @param tagName the element's lowercase tag name
   * @param parent the (possibly simulated) parent node
   * @param parentTagName the parent's lowercase tag name
   * @returns true if a spec-compliant parser could produce this element
   */
  const _checkMathMlNamespace = function _checkMathMlNamespace(tagName, parent, parentTagName) {
    // The only way to switch from HTML namespace to MathML
    // is via <math>. If it happens via any other tag, then
    // it should be killed.
    if (parent.namespaceURI === HTML_NAMESPACE) {
      return tagName === 'math';
    }
    // The only way to switch from SVG to MathML is via
    // <math> and HTML integration points
    if (parent.namespaceURI === SVG_NAMESPACE) {
      return tagName === 'math' && HTML_INTEGRATION_POINTS[parentTagName];
    }
    // We only allow elements that are defined in MathML
    // spec. All others are disallowed in MathML namespace.
    return Boolean(ALL_MATHML_TAGS[tagName]);
  };
  /**
   * Namespace rules for an element in the HTML namespace.
   *
   * @param tagName the element's lowercase tag name
   * @param parent the (possibly simulated) parent node
   * @param parentTagName the parent's lowercase tag name
   * @returns true if a spec-compliant parser could produce this element
   */
  const _checkHtmlNamespace = function _checkHtmlNamespace(tagName, parent, parentTagName) {
    // The only way to switch from SVG to HTML is via
    // HTML integration points, and from MathML to HTML
    // is via MathML text integration points
    if (parent.namespaceURI === SVG_NAMESPACE && !HTML_INTEGRATION_POINTS[parentTagName]) {
      return false;
    }
    if (parent.namespaceURI === MATHML_NAMESPACE && !MATHML_TEXT_INTEGRATION_POINTS[parentTagName]) {
      return false;
    }
    // We disallow tags that are specific for MathML
    // or SVG and should never appear in HTML namespace
    return !ALL_MATHML_TAGS[tagName] && (COMMON_SVG_AND_HTML_ELEMENTS[tagName] || !ALL_SVG_TAGS[tagName]);
  };
  /**
   * @param element a DOM element whose namespace is being checked
   * @returns Return false if the element has a
   *  namespace that a spec-compliant parser would never
   *  return. Return true otherwise.
   */
  const _checkValidNamespace = function _checkValidNamespace(element) {
    let parent = getParentNode(element);
    // In JSDOM, if we're inside shadow DOM, then parentNode
    // can be null. We just simulate parent in this case.
    if (!parent || !parent.tagName) {
      parent = {
        namespaceURI: NAMESPACE,
        tagName: 'template'
      };
    }
    const tagName = stringToLowerCase(element.tagName);
    const parentTagName = stringToLowerCase(parent.tagName);
    if (!ALLOWED_NAMESPACES[element.namespaceURI]) {
      return false;
    }
    if (element.namespaceURI === SVG_NAMESPACE) {
      return _checkSvgNamespace(tagName, parent, parentTagName);
    }
    if (element.namespaceURI === MATHML_NAMESPACE) {
      return _checkMathMlNamespace(tagName, parent, parentTagName);
    }
    if (element.namespaceURI === HTML_NAMESPACE) {
      return _checkHtmlNamespace(tagName, parent, parentTagName);
    }
    // For XHTML and XML documents that support custom namespaces
    if (PARSER_MEDIA_TYPE === 'application/xhtml+xml' && ALLOWED_NAMESPACES[element.namespaceURI]) {
      return true;
    }
    // The code should never reach this place (this means
    // that the element somehow got namespace that is not
    // HTML, SVG, MathML or allowed via ALLOWED_NAMESPACES).
    // Return false just in case.
    return false;
  };
  /**
   * _forceRemove
   *
   * @param node a DOM node
   */
  const _forceRemove = function _forceRemove(node) {
    arrayPush(DOMPurify.removed, {
      element: node
    });
    try {
      // eslint-disable-next-line unicorn/prefer-dom-node-remove
      getParentNode(node).removeChild(node);
    } catch (_) {
      /* The normal detach failed — this is reached for a parentless node
         (getParentNode() is null, so .removeChild throws). Element.prototype
         .remove() is itself a spec no-op on a parentless node, so a recorded
         "removal" would otherwise hand the caller back an intact,
         payload-bearing node (e.g. a detached IN_PLACE root the mXSS canary or
         the style-with-element-child rule decided to kill). Fail closed by
         throwing — exactly as a clobbered root does at the IN_PLACE entry —
         rather than trying to "neutralize" the node via its own methods.
         Neutralizing would mean calling getAttributeNames()/removeAttribute()
         on the node, both of which a <form> root can clobber via a named child
         (and _isClobbered does not even probe getAttributeNames), so the
         neutralize step could itself be silently defeated, leaving the payload
         intact. A throw touches only the cached, clobber-safe remove() and
         getParentNode(). Generalizes GHSA-r47g-fvhr-h676 (clobbered-form root)
         to every root-kill reason. REPORT-3.
                This lives inside the catch, so it never fires for a normally-removed
         in-tree node: those have a parent, removeChild() succeeds, and the
         catch is not entered. Only a kept (parentless) root reaches here. */
      remove(node);
      if (!getParentNode(node)) {
        throw typeErrorCreate('a node selected for removal could not be detached from its tree ' + 'and cannot be safely returned; refusing to sanitize in place');
      }
    }
  };
  /**
   * _neutralizeRoot
   *
   * Fail-closed teardown of an in-place root after the sanitize walk aborts
   * (campaign-3 F2). An internal throw mid-walk — e.g. a page-registered
   * custom element's reaction detaches a node so `_forceRemove`'s deliberate
   * parentless guard throws, or any other re-entrant engine mutation — would
   * otherwise leave the caller's *live* tree half-sanitized, with everything
   * after the abort point still carrying its handlers. There is no safe way
   * to resume the walk (the tree mutated under us), so we strip the root bare:
   * remove every child and every attribute, then let the caller's catch see
   * the original error. Clobber-safe (cached `remove`/`childNodes`/`attributes`
   * getters; the root was already clobber-pre-flighted at the IN_PLACE entry).
   *
   * @param root the in-place root to empty
   */
  const _neutralizeRoot = function _neutralizeRoot(root) {
    /* Strip every disallowed attribute (on* handlers included) off the whole
       subtree BEFORE detaching anything. Detaching first would hand back
       handler-bearing originals (e.g. an already-loading `<img onerror>`)
       whose queued resource event still fires in page scope after we throw.
       Clobber-safe reads; a doomed clobbered node's own attributes are
       irrelevant while its non-clobbered descendants are reached and scrubbed. */
    _neutralizeSubtree(root);
    const childNodes = getChildNodes(root);
    if (childNodes) {
      const snapshot = [];
      arrayForEach(childNodes, child => {
        arrayPush(snapshot, child);
      });
      arrayForEach(snapshot, child => {
        try {
          remove(child);
        } catch (_) {
          /* Best-effort teardown; a still-attached child is handled below */
        }
      });
    }
    const attributes = getAttributes(root);
    if (attributes) {
      for (let i = attributes.length - 1; i >= 0; --i) {
        const attribute = attributes[i];
        const name = attribute && attribute.name;
        if (typeof name === 'string') {
          try {
            root.removeAttribute(name);
          } catch (_) {
            /* Clobbered removeAttribute — ignore (fail-closed best effort) */
          }
        }
      }
    }
  };
  /**
   * _removeAttribute
   *
   * @param name an Attribute name
   * @param element a DOM node
   */
  const _removeAttribute = function _removeAttribute(name, element) {
    try {
      arrayPush(DOMPurify.removed, {
        attribute: element.getAttributeNode(name),
        from: element
      });
    } catch (_) {
      arrayPush(DOMPurify.removed, {
        attribute: null,
        from: element
      });
    }
    element.removeAttribute(name);
    // We void attribute values for unremovable "is" attributes
    if (name === 'is') {
      if (RETURN_DOM || RETURN_DOM_FRAGMENT) {
        try {
          _forceRemove(element);
        } catch (_) {}
      } else {
        try {
          element.setAttribute(name, '');
        } catch (_) {}
      }
    }
  };
  /**
   * _stripDisallowedAttributes
   *
   * Removes every attribute the active configuration does not allow from a
   * single element, using the same allowlist as the main attribute pass (so
   * `on*` handlers go, but no `/^on/` blocklist is introduced). Used only to
   * neutralise nodes that are being discarded from an in-place tree.
   *
   * @param element the element to strip
   */
  const _stripDisallowedAttributes = function _stripDisallowedAttributes(element) {
    const attributes = getAttributes(element);
    if (!attributes) {
      return;
    }
    for (let i = attributes.length - 1; i >= 0; --i) {
      const attribute = attributes[i];
      const name = attribute && attribute.name;
      if (typeof name !== 'string' || ALLOWED_ATTR[transformCaseFunc(name)]) {
        continue;
      }
      try {
        element.removeAttribute(name);
      } catch (_) {
        /* Clobbered removeAttribute on a doomed node — ignore */
      }
    }
  };
  /**
   * _neutralizeSubtree
   *
   * Completes the audit-5 F1 fix across every removal path. The KEEP_CONTENT
   * move-hoist neutralises only disallowed-tag removals; clobber, mXSS-canary,
   * namespace, comment, processing-instruction and KEEP_CONTENT:false removals
   * all drop their subtree wholesale via `_forceRemove`. On the IN_PLACE path
   * those dropped nodes are detached from the caller's LIVE tree but a
   * handler-bearing original among them (an `<img onerror>`/`<video>` that was
   * loading) keeps its queued resource event, which fires in page scope after
   * sanitize returns. This walks a removed subtree and strips every attribute
   * the active configuration does not allow — so `on*` handlers are cancelled
   * through the SAME allowlist that governs kept nodes, not a separate `/^on/`
   * blocklist. Run synchronously before sanitize returns, i.e. before any
   * queued event can fire. Hook-free by design: these nodes leave the output,
   * so firing attribute hooks for them would be surprising. Clobber-safe reads;
   * a doomed clobbered node may shadow `removeAttribute` (its own attributes are
   * irrelevant — it is discarded — while its non-clobbered descendants, e.g.
   * the `<img>`, are reached and scrubbed).
   *
   * @param root the root of a removed subtree to neutralise
   */
  const _neutralizeSubtree = function _neutralizeSubtree(root) {
    const stack = [root];
    while (stack.length > 0) {
      const node = stack.pop();
      const nodeType = getNodeType ? getNodeType(node) : node.nodeType;
      if (nodeType === NODE_TYPE.element) {
        _stripDisallowedAttributes(node);
      }
      const childNodes = getChildNodes(node);
      if (childNodes) {
        for (let i = childNodes.length - 1; i >= 0; --i) {
          stack.push(childNodes[i]);
        }
      }
    }
  };
  /**
   * _neutralizePatchLinkage
   *
   * IN_PLACE entry pre-pass (declarative-partial-updates / streaming
   * hardening, https://github.com/WICG/declarative-partial-updates).
   *
   * The main walk strips patch linkage (`for`/`patchsrc`) and removes range
   * markers (PIs / markup comments) node-by-node, in document order, AS it
   * reaches each node. On a live in-place root that leaves a window: from the
   * moment the root is connected until the walk arrives at a given node, that
   * node's linkage is live. A patch applied on connection/stream can fire as
   * a microtask during the walk and inject or teleport an unsanitized DOM
   * range into a region the iterator has already passed and will not revisit,
   * so the post-return "tree is sanitized" contract is violated. Sweep the
   * whole tree once up front and sever every linkage before the walk begins,
   * closing that window.
   *
   * This CANNOT undo a patch that already fired before sanitize ran — that is
   * the irreducible "do not IN_PLACE a live-connected attacker tree" caveat —
   * but it closes everything from sanitize-start onward. Gated on SAFE_FOR_XML
   * to group with the rest of the declarative-partial-updates handling and
   * stay overridable, consistent with the codebase.
   *
   * Clobber-safe traversal (cached childNodes getter); per-node try/catch so a
   * clobbered root cannot defeat the sweep of its non-clobbered descendants.
   *
   * NOTE (pending real-Chrome confirmation, see test/declarative-patch-probe
   * .html Q1): this mirrors the existing policy of keeping `for` on
   * <label>/<output>. If the shipping feature can drive a patch through a
   * surviving `for`-on-label/output + `id` pair, this pre-pass and the
   * attribute check at _isBasicCustomElement's caller must additionally drop
   * that pair on the IN_PLACE path. Left as-is until the taxonomy is verified.
   *
   * @param root the in-place root to sweep
   */
  const _neutralizePatchLinkage = function _neutralizePatchLinkage(root) {
    if (!SAFE_FOR_XML) {
      return;
    }
    const stack = [root];
    while (stack.length > 0) {
      const node = stack.pop();
      const nodeType = getNodeType ? getNodeType(node) : node.nodeType;
      /* Remove range markers (the target side of a patch linkage): every
         processing instruction, and any markup-bearing comment. */
      if (nodeType === NODE_TYPE.processingInstruction || nodeType === NODE_TYPE.comment && regExpTest(COMMENT_MARKUP_PROBE, node.data)) {
        try {
          remove(node);
        } catch (_) {
          /* Best-effort */
        }
        continue;
      }
      /* Strip patch-source attributes (the source side) off elements. */
      if (nodeType === NODE_TYPE.element) {
        const element = node;
        const lcTag = transformCaseFunc(getNodeName ? getNodeName(node) : node.nodeName);
        try {
          if (element.hasAttribute && element.hasAttribute('patchsrc')) {
            element.removeAttribute('patchsrc');
          }
          if (element.hasAttribute && element.hasAttribute('for') && lcTag !== 'label' && lcTag !== 'output') {
            element.removeAttribute('for');
          }
        } catch (_) {
          /* Clobbered removeAttribute/hasAttribute on a doomed node — ignore */
        }
      }
      const childNodes = getChildNodes(node);
      if (childNodes) {
        for (let i = childNodes.length - 1; i >= 0; --i) {
          stack.push(childNodes[i]);
        }
      }
    }
  };
  /**
   * _initDocument
   *
   * @param dirty - a string of dirty markup
   * @return a DOM, filled with the dirty markup
   */
  const _initDocument = function _initDocument(dirty) {
    /* Create a HTML document */
    let doc = null;
    let leadingWhitespace = null;
    if (FORCE_BODY) {
      dirty = '<remove></remove>' + dirty;
    } else {
      /* If FORCE_BODY isn't used, leading whitespace needs to be preserved manually */
      const matches = stringMatch(dirty, /^[\r\n\t ]+/);
      leadingWhitespace = matches && matches[0];
    }
    if (PARSER_MEDIA_TYPE === 'application/xhtml+xml' && NAMESPACE === HTML_NAMESPACE) {
      // Root of XHTML doc must contain xmlns declaration (see https://www.w3.org/TR/xhtml1/normative.html#strict)
      dirty = '<html xmlns="http://www.w3.org/1999/xhtml"><head></head><body>' + dirty + '</body></html>';
    }
    const dirtyPayload = trustedTypesPolicy ? _createTrustedHTML(dirty) : dirty;
    /*
     * Use the DOMParser API by default, fallback later if needs be
     * DOMParser not work for svg when has multiple root element.
     */
    if (NAMESPACE === HTML_NAMESPACE) {
      try {
        doc = new DOMParser().parseFromString(dirtyPayload, PARSER_MEDIA_TYPE);
      } catch (_) {}
    }
    /* Use createHTMLDocument in case DOMParser is not available */
    if (!doc || !doc.documentElement) {
      doc = implementation.createDocument(NAMESPACE, 'template', null);
      try {
        doc.documentElement.innerHTML = IS_EMPTY_INPUT ? emptyHTML : dirtyPayload;
      } catch (_) {
        // Syntax error if dirtyPayload is invalid xml
      }
    }
    const body = doc.body || doc.documentElement;
    if (dirty && leadingWhitespace) {
      body.insertBefore(document.createTextNode(leadingWhitespace), body.childNodes[0] || null);
    }
    /* Work on whole document or just its body */
    if (NAMESPACE === HTML_NAMESPACE) {
      return getElementsByTagName.call(doc, WHOLE_DOCUMENT ? 'html' : 'body')[0];
    }
    return WHOLE_DOCUMENT ? doc.documentElement : body;
  };
  /**
   * Creates a NodeIterator object that you can use to traverse filtered lists of nodes or elements in a document.
   *
   * @param root The root element or node to start traversing on.
   * @return The created NodeIterator
   */
  const _createNodeIterator = function _createNodeIterator(root) {
    /* Read ownerDocument through the cached Node.prototype getter, never the
       direct property. HTMLFormElement has [LegacyOverrideBuiltIns], so a
       clobbering child (<input name="ownerDocument"> or a form-associated
       external input) shadows the prototype getter and makes a direct read
       return that <input>. createNodeIterator.call(<input>, ...) then throws
       "Illegal invocation", and on the IN_PLACE path that throw lands before
       the walk's fail-closed barrier - leaving the caller's live tree, with
       any already-armed handler in it, un-neutralized. The cached getter
       returns the real Document regardless of the clobber. */
    const doc = getOwnerDocument ? getOwnerDocument(root) : root.ownerDocument;
    return createNodeIterator.call(doc || root, root,
    // eslint-disable-next-line no-bitwise
    NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_COMMENT | NodeFilter.SHOW_TEXT | NodeFilter.SHOW_PROCESSING_INSTRUCTION | NodeFilter.SHOW_CDATA_SECTION, null);
  };
  /**
   * Replace template expression syntax (mustache, ERB, template
   * literal) with a space; shared by all SAFE_FOR_TEMPLATES scrub
   * sites. Order matters: mustache, then ERB, then template literal.
   *
   * @param value the string to scrub
   * @returns the scrubbed string
   */
  const _stripTemplateExpressions = function _stripTemplateExpressions(value) {
    value = stringReplace(value, MUSTACHE_EXPR$1, ' ');
    value = stringReplace(value, ERB_EXPR$1, ' ');
    value = stringReplace(value, TMPLIT_EXPR$1, ' ');
    return value;
  };
  /**
   * Strip template-engine expressions ({{...}}, ${...}, <%...%>) from the
   * character data of an element subtree. Used as the final safety net for
   * SAFE_FOR_TEMPLATES on every DOM-returning code path so that expressions
   * which only form after text-node normalization (e.g. fragments split across
   * stripped elements) cannot survive into a template-evaluating framework.
   *
   * Walks text/comment/CDATA/processing-instruction nodes and mutates `.data`
   * in place rather than round-tripping through innerHTML. This preserves
   * descendant node references (important for IN_PLACE callers), avoids a
   * serialize/reparse cycle, and reads literal character data — which means
   * `<%...%>` in text content matches the ERB regex against its real bytes
   * instead of the HTML-entity-escaped form innerHTML would produce.
   *
   * Attribute values are not visited here; SAFE_FOR_TEMPLATES handling for
   * attributes is performed during the per-node `_sanitizeAttributes` pass.
   *
   * @param node The root element whose character data should be scrubbed.
   */
  const _scrubTemplateExpressions2 = function _scrubTemplateExpressions(node) {
    var _node$querySelectorAl;
    node.normalize();
    /* Clobber-safe ownerDocument read, same reasoning as _createNodeIterator:
       under SAFE_FOR_TEMPLATES this runs on the live IN_PLACE root, which may
       carry a form-named-getter override of ownerDocument. */
    const doc = getOwnerDocument ? getOwnerDocument(node) : node.ownerDocument;
    const walker = createNodeIterator.call(doc || node, node,
    // eslint-disable-next-line no-bitwise
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_COMMENT | NodeFilter.SHOW_CDATA_SECTION | NodeFilter.SHOW_PROCESSING_INSTRUCTION, null);
    let currentNode = walker.nextNode();
    while (currentNode) {
      currentNode.data = _stripTemplateExpressions(currentNode.data);
      currentNode = walker.nextNode();
    }
    // NodeIterator does not descend into <template>.content per the DOM spec,
    // so we must explicitly recurse into each template's content fragment,
    // mirroring the approach used by _sanitizeShadowDOM.
    const templates = (_node$querySelectorAl = node.querySelectorAll) === null || _node$querySelectorAl === void 0 ? void 0 : _node$querySelectorAl.call(node, 'template');
    if (templates) {
      arrayForEach(templates, tmpl => {
        if (_isDocumentFragment(tmpl.content)) {
          _scrubTemplateExpressions2(tmpl.content);
        }
      });
    }
  };
  /**
   * _isClobbered
   *
   * Detect DOM-clobbering on HTMLFormElement nodes. Form is the only HTML
   * interface with [LegacyOverrideBuiltIns]; a descendant element with a
   * `name` attribute matching a prototype property shadows that property
   * on direct reads. We use this check at the IN_PLACE entry-point and
   * during attribute sanitization to refuse clobbered forms.
   *
   * @param element element to check for clobbering attacks
   * @return true if clobbered, false if safe
   */
  const _isClobbered = function _isClobbered(element) {
    // Realm-independent tag-name probe. If we can't determine the tag
    // name at all, we can't reason about clobbering — return false
    // (the caller's other defences still apply).
    const realTagName = getNodeName ? getNodeName(element) : null;
    if (typeof realTagName !== 'string') {
      return false;
    }
    if (transformCaseFunc(realTagName) !== 'form') {
      return false;
    }
    return typeof element.nodeName !== 'string' || typeof element.textContent !== 'string' || typeof element.removeChild !== 'function' ||
    // Realm-safe NamedNodeMap detection: equality against the cached
    // prototype getter. Clobbered .attributes (e.g. <input name="attributes">)
    // makes the direct read diverge from the cached read; a clean form
    // (same-realm OR foreign-realm) has both reads pointing at the same
    // canonical NamedNodeMap.
    element.attributes !== getAttributes(element) || typeof element.removeAttribute !== 'function' || typeof element.setAttribute !== 'function' || typeof element.namespaceURI !== 'string' || typeof element.insertBefore !== 'function' || typeof element.hasChildNodes !== 'function' ||
    // NodeType clobbering probe. Cached Node.prototype.nodeType getter
    // returns the integer 1 for any Element regardless of realm; direct
    // read on a clobbered form (e.g. <input name="nodeType">) returns
    // the named child element. Cheap addition — nodeType is read from
    // an internal slot, no serialization cost — and removes a residual
    // clobbering surface used by several mXSS / PI / comment branches
    // in _sanitizeElements that compare currentNode.nodeType directly.
    element.nodeType !== getNodeType(element) ||
    // HTMLFormElement has [LegacyOverrideBuiltIns]: a descendant named
    // "childNodes" shadows the prototype getter. Direct reads of
    // form.childNodes from a clobbered form return the named child
    // instead of the real NodeList, so any walk that reads it directly
    // skips the form's real children. Compare the direct read to the
    // cached Node.prototype getter — when the form's named-property
    // getter intercepts the read, the two values differ and we flag
    // the form. This catches every clobbering child type (input,
    // select, etc.) regardless of whether the named child happens to
    // carry a numeric .length, which a typeof-based probe would miss
    // (e.g. HTMLSelectElement.length is a defined unsigned-long).
    element.childNodes !== getChildNodes(element);
  };
  /**
   * Checks whether the given value is a DocumentFragment from any realm.
   *
   * The realm-independent replacement reads `nodeType` through the cached
   * Node.prototype getter and compares to the DOCUMENT_FRAGMENT_NODE
   * constant (11). nodeType is a numeric value resolved from the node's
   * internal slot, identical across realms for the same kind of node.
   *
   * @param value object to check
   * @return true if value is a DocumentFragment-shaped node from any realm
   */
  const _isDocumentFragment = function _isDocumentFragment(value) {
    if (!getNodeType || typeof value !== 'object' || value === null) {
      return false;
    }
    try {
      return getNodeType(value) === NODE_TYPE.documentFragment;
    } catch (_) {
      return false;
    }
  };
  /**
   * Checks whether the given object is a DOM node, including nodes that
   * originate from a different window/realm (e.g. an iframe's
   * contentDocument). The previous `value instanceof Node` check was
   * realm-bound: nodes from a different window failed it, causing
   * sanitize() to silently stringify them and reset IN_PLACE to false,
   * returning the original node unsanitized. See GHSA-4w3q-35jp-p934.
   *
   * @param value object to check whether it's a DOM node
   * @return true if value is a DOM node from any realm
   */
  const _isNode = function _isNode(value) {
    if (!getNodeType || typeof value !== 'object' || value === null) {
      return false;
    }
    try {
      return typeof getNodeType(value) === 'number';
    } catch (_) {
      return false;
    }
  };
  function _executeHooks(hooks, currentNode, data) {
    if (hooks.length === 0) {
      return;
    }
    arrayForEach(hooks, hook => {
      hook.call(DOMPurify, currentNode, data, CONFIG);
    });
  }
  /**
   * Structural-threat checks that condemn a node regardless of the
   * allowlists: mXSS via namespace confusion, risky CSS construction,
   * processing instructions, markup-bearing comments. Pure predicate;
   * the caller removes. Check order is load-bearing.
   *
   * @param currentNode the node to inspect
   * @param tagName the node's transformCaseFunc'd tag name
   * @return true if the node must be removed
   */
  const _isUnsafeNode = function _isUnsafeNode(currentNode, tagName) {
    /* Detect mXSS attempts abusing namespace confusion */
    if (SAFE_FOR_XML && currentNode.hasChildNodes() && !_isNode(currentNode.firstElementChild) && regExpTest(ELEMENT_MARKUP_PROBE, currentNode.textContent) && regExpTest(ELEMENT_MARKUP_PROBE, currentNode.innerHTML)) {
      return true;
    }
    /* Remove risky CSS construction leading to mXSS */
    if (SAFE_FOR_XML && currentNode.namespaceURI === HTML_NAMESPACE && tagName === 'style' && _isNode(currentNode.firstElementChild)) {
      return true;
    }
    /* Remove any occurrence of processing instructions */
    if (currentNode.nodeType === NODE_TYPE.processingInstruction) {
      return true;
    }
    /* Remove any kind of possibly harmful comments */
    if (SAFE_FOR_XML && currentNode.nodeType === NODE_TYPE.comment && regExpTest(COMMENT_MARKUP_PROBE, currentNode.data)) {
      return true;
    }
    return false;
  };
  /**
   * Handle a node whose tag is forbidden or not allowlisted: keep
   * allowed custom elements (false return exits _sanitizeElements
   * early - the namespace and fallback-tag removal checks are
   * intentionally skipped for kept custom elements), else hoist
   * content per KEEP_CONTENT and remove.
   *
   * A kept custom element is the ONLY case in which this function
   * returns false, so the caller uses that return value to run the
   * afterSanitizeElements hook on the kept element and keep the
   * element-hook lifecycle consistent with normal allowlisted
   * elements (GHSA-c2j3-45gr-mqc4).
   *
   * @param currentNode the disallowed node
   * @param tagName the node's transformCaseFunc'd tag name
   * @return true if the node was removed, false if kept
   */
  const _sanitizeDisallowedNode = function _sanitizeDisallowedNode(currentNode, tagName, root) {
    /* Check if we have a custom element to handle */
    if (!FORBID_TAGS[tagName] && _isBasicCustomElement(tagName)) {
      if (CUSTOM_ELEMENT_HANDLING.tagNameCheck instanceof RegExp && regExpTest(CUSTOM_ELEMENT_HANDLING.tagNameCheck, tagName)) {
        return false;
      }
      if (CUSTOM_ELEMENT_HANDLING.tagNameCheck instanceof Function && CUSTOM_ELEMENT_HANDLING.tagNameCheck(tagName)) {
        return false;
      }
    }
    /* Keep content except for bad-listed elements.
         Use the cached prototype getters exclusively — the previous code
         had `|| currentNode.parentNode` / `|| currentNode.childNodes`
         fallbacks, but the cached getters always return the canonical
         value (or null for a real parent-less node), so the fallback
         path was dead in safe cases and a clobbering surface in unsafe
         ones. Falsy cached results stay falsy; the `if (childNodes &&
         parentNode)` check already gates correctly. */
    if (KEEP_CONTENT && !FORBID_CONTENTS[tagName]) {
      const parentNode = getParentNode(currentNode);
      const childNodes = getChildNodes(currentNode);
      if (childNodes && parentNode) {
        const childCount = childNodes.length;
        /* Hoist by moving each child up one level rather than deep-cloning
             it. Moving transfers every descendant exactly once, so a chain of
             nested disallowed elements costs O(n) instead of the O(n^2) that
             re-cloning the shrinking subtree at each level produced; it also
             empties the removed original, so `DOMPurify.removed` no longer
             pins whole subtrees. Moving preserves the in-place guarantee too:
             an original carrying already-queued resource events (`<img
             onerror>`, `<video>`/`<audio>` error, lazy/`onload`, …) is
             relocated and sanitised rather than left detached but still armed.
                      The sole case that must clone is removing the walk root itself.
             The result is serialised from the root's subtree, so a restrictive
             ALLOWED_TAGS that strips the root (`body` on the string path) must
             leave the content inside it, which only cloning does. In IN_PLACE
             the root is pre-validated as an allowed tag and so is never removed
             here, so that path always takes the move branch.
                      `childNodes` is live; a tail-to-head walk keeps `childNodes[i]`
             valid whether we move (drops the trailing entry) or clone (leaves
             the list intact). */
        for (let i = childCount - 1; i >= 0; --i) {
          const hoisted = currentNode === root ? cloneNode(childNodes[i], true) : childNodes[i];
          parentNode.insertBefore(hoisted, getNextSibling(currentNode));
        }
      }
    }
    _forceRemove(currentNode);
    return true;
  };
  /**
   * Fork a hook-mutable allowlist off its shared binding the first time a
   * (possibly lazily-installed) uponSanitize* hook is about to see it, so the
   * hook cannot widen the per-instance default or the setConfig binding by
   * reference and leak past the call. Returns the set unchanged once it is
   * already call-local, so repeated calls across elements are idempotent.
   *
   * @param hookList the uponSanitize* hook array for this event
   * @param set the current ALLOWED_TAGS / ALLOWED_ATTR binding
   * @param defaultSet the per-instance DEFAULT_ALLOWED_* constant
   * @param setConfigSet the captured setConfig() binding, or null
   * @return a call-local clone if a hook is present and set is still shared,
   *   else set unchanged
   */
  const _forkSharedAllowlist = function _forkSharedAllowlist(hookList, set, defaultSet, setConfigSet) {
    if (hookList.length === 0) {
      return set;
    }
    return set === defaultSet || set === setConfigSet ? clone(set) : set;
  };
  /**
   * _sanitizeElements
   *
   * @protect nodeName
   * @protect textContent
   * @protect removeChild
   * @param currentNode to check for permission to exist
   * @return true if node was killed, false if left alive
   */
  // eslint-disable-next-line complexity
  const _sanitizeElements = function _sanitizeElements(currentNode, root) {
    /* Execute a hook if present */
    _executeHooks(hooks.beforeSanitizeElements, currentNode, null);
    /* A hook may have detached the node - treat it as removed (see the
       detached-node comment after the uponSanitizeElement hook below). On
       the IN_PLACE path, neutralize the detached subtree first so a queued
       resource handler on it cannot fire in page scope after we return. */
    if (currentNode !== root && getParentNode(currentNode) === null) {
      if (IN_PLACE) {
        _neutralizeSubtree(currentNode);
      }
      return true;
    }
    /* Check if element is clobbered or can clobber */
    if (_isClobbered(currentNode)) {
      _forceRemove(currentNode);
      return true;
    }
    /* Now let's check the element's type and name */
    const tagName = transformCaseFunc(getNodeName ? getNodeName(currentNode) : currentNode.nodeName);
    /* Close the pre-walk clone-guard's timing gap: an uponSanitizeElement
       hook may have been installed after that guard sampled the hook arrays
       (e.g. lazily from beforeSanitizeElements), leaving ALLOWED_TAGS still
       aliasing a shared binding that a widening hook would mutate by
       reference. Fork it before exposing it to the hook. */
    ALLOWED_TAGS = _forkSharedAllowlist(hooks.uponSanitizeElement, ALLOWED_TAGS, DEFAULT_ALLOWED_TAGS, SET_CONFIG_ALLOWED_TAGS);
    /* Execute a hook if present */
    _executeHooks(hooks.uponSanitizeElement, currentNode, {
      tagName,
      allowedTags: ALLOWED_TAGS
    });
    /* A hook may have detached the node from the tree — a long-standing
       user pattern (issue #469; draw.io-style foreignObject filtering).
       Per the cached, unclobberable parentNode getter the node is
       genuinely out of the tree, so it can reach neither the serialized
       output nor an IN_PLACE live tree; treat it as removed and stop
       processing it. Without this guard, the unsafe-node / namespace
       checks below would call _forceRemove on a parentless node and hit
       the REPORT-3 fail-closed throw — which exists for nodes DOMPurify
       wants gone but *cannot* detach (clobbered / parentless roots), the
       opposite of a node that is already safely gone. The walk root is
       exempt: a detached IN_PLACE root is legitimate input and must still
       be fully sanitized, and a kill-decision on it must keep hitting the
       REPORT-3 throw. Nodes detached by hooks stay the hook's
       responsibility for placement: they are not recorded in
       DOMPurify.removed, so the post-walk IN_PLACE pass (which iterates
       DOMPurify.removed) does not reach them. But a hook-detached subtree
       can still hold a queued resource-event handler - e.g. an <img onload>
       that began loading when the caller built the live tree - which fires
       in page scope after sanitize returns even though the handler never
       reached the returned tree. That is the audit-5 F1 hazard, and the
       documented node.remove() hook pattern walks straight into it. So on
       the IN_PLACE path we neutralize the detached subtree inline here,
       stripping its non-allow-listed attributes before returning, exactly
       as the post-walk pass does for _forceRemove'd subtrees. */
    if (currentNode !== root && getParentNode(currentNode) === null) {
      if (IN_PLACE) {
        _neutralizeSubtree(currentNode);
      }
      return true;
    }
    /* Remove mXSS vectors, processing instructions and risky comments */
    if (_isUnsafeNode(currentNode, tagName)) {
      _forceRemove(currentNode);
      return true;
    }
    /* Remove element if anything forbids its presence */
    if (FORBID_TAGS[tagName] || !(EXTRA_ELEMENT_HANDLING.tagCheck instanceof Function && EXTRA_ELEMENT_HANDLING.tagCheck(tagName)) && !ALLOWED_TAGS[tagName]) {
      const removed = _sanitizeDisallowedNode(currentNode, tagName, root);
      /* A false return means the node is a custom element kept via
         CUSTOM_ELEMENT_HANDLING - the only keep path through
         _sanitizeDisallowedNode. Run afterSanitizeElements on it so the
         element-hook lifecycle matches normal allowlisted elements: a
         security policy applied in this hook (e.g. stripping an attribute
         from every surviving element) must not silently skip kept custom
         elements (GHSA-c2j3-45gr-mqc4). This mirrors the normal-element
         tail below - the hook runs, then the walker's subsequent
         _sanitizeAttributes pass sanitizes the element's attributes. The
         deliberately skipped namespace and fallback-tag removal checks stay
         skipped; they are removal decisions, not the hook contract. */
      if (removed === false) {
        _executeHooks(hooks.afterSanitizeElements, currentNode, null);
      }
      return removed;
    }
    /* Check whether element has a valid namespace.
       Realm-safe check (GHSA-hpcv-96wg-7vj8): use the cached Node.prototype
       nodeType getter rather than `instanceof Element`, which is realm-
       bound and short-circuits to false for any node minted in a different
       realm — letting a foreign-realm element with a forbidden namespace
       slip past the namespace check entirely. */
    const nt = getNodeType ? getNodeType(currentNode) : currentNode.nodeType;
    if (nt === NODE_TYPE.element && !_checkValidNamespace(currentNode)) {
      _forceRemove(currentNode);
      return true;
    }
    /* Make sure that older browsers don't get fallback-tag mXSS */
    if ((tagName === 'noscript' || tagName === 'noembed' || tagName === 'noframes') && regExpTest(FALLBACK_TAG_CLOSE, currentNode.innerHTML)) {
      _forceRemove(currentNode);
      return true;
    }
    /* Sanitize element content to be template-safe */
    if (SAFE_FOR_TEMPLATES && currentNode.nodeType === NODE_TYPE.text) {
      /* Get the element's text content */
      const content = _stripTemplateExpressions(currentNode.textContent);
      if (currentNode.textContent !== content) {
        arrayPush(DOMPurify.removed, {
          element: currentNode.cloneNode()
        });
        currentNode.textContent = content;
      }
    }
    /* Execute a hook if present */
    _executeHooks(hooks.afterSanitizeElements, currentNode, null);
    return false;
  };
  /**
   * _isValidAttribute
   *
   * @param lcTag Lowercase tag name of containing element.
   * @param lcName Lowercase attribute name.
   * @param value Attribute value.
   * @return Returns true if `value` is valid, otherwise false.
   */
  // eslint-disable-next-line complexity
  const _isValidAttribute = function _isValidAttribute(lcTag, lcName, value) {
    /* FORBID_ATTR must always win, even if ADD_ATTR predicate would allow it */
    if (FORBID_ATTR[lcName]) {
      return false;
    }
    /* Reject declarative-partial-updates patch-linkage attributes
       (https://github.com/WICG/declarative-partial-updates).
            Empirical note (Chrome 150, verified — see
       test/declarative-patch-probe-v3.html): expansion is NOT applied after
       sanitization. For the string path it fires during sanitize()'s own
       parse, so the walk sees and sanitizes the fully materialized expanded
       tree — teleports into MathML/SVG integration points included; a
       weaponized `<template for>`->`<img onerror>` comes back with the handler
       stripped. For the IN_PLACE path it fires on connection, before the walk.
       Either way DOMPurify is NOT blind to the patch.
            This removal is therefore defense-in-depth rather than the sole barrier:
       it prevents live linkage from surviving into the OUTPUT and re-expanding
       in the caller's context, and keeps behaviour deterministic if a future
       engine defers expansion. `for` is legitimate only on <label>/<output>;
       anywhere else (notably <template for>) it links the element to a patch
       target and teleports or removes an arbitrary DOM range by id/marker name.
       `patchsrc` fetches remote markup and is treated as a script-loading
       mechanism (CSP). Gated on SAFE_FOR_XML so the removal groups with the
       other structural-threat checks and stays overridable, consistent with
       the rest of the codebase. PI range markers are already removed by
       _isUnsafeNode. */
    if (SAFE_FOR_XML && lcName === 'patchsrc') {
      return false;
    }
    if (SAFE_FOR_XML && lcName === 'for' && lcTag !== 'label' && lcTag !== 'output') {
      return false;
    }
    /* Make sure attribute cannot clobber */
    if (SANITIZE_DOM && (lcName === 'id' || lcName === 'name') && (value in document || value in formElement)) {
      return false;
    }
    const nameIsPermitted = ALLOWED_ATTR[lcName] || EXTRA_ELEMENT_HANDLING.attributeCheck instanceof Function && EXTRA_ELEMENT_HANDLING.attributeCheck(lcName, lcTag);
    /* Allow valid data-* attributes: At least one character after "-"
        (https://html.spec.whatwg.org/multipage/dom.html#embedding-custom-non-visible-data-with-the-data-*-attributes)
        XML-compatible (https://html.spec.whatwg.org/multipage/infrastructure.html#xml-compatible and http://www.w3.org/TR/xml/#d0e804)
        We don't need to check the value; it's always URI safe. */
    if (ALLOW_DATA_ATTR && regExpTest(DATA_ATTR$1, lcName)) ; else if (ALLOW_ARIA_ATTR && regExpTest(ARIA_ATTR$1, lcName)) ; else if (!nameIsPermitted) {
      if (
      // First condition does a very basic check if a) it's basically a valid custom element tagname AND
      // b) if the tagName passes whatever the user has configured for CUSTOM_ELEMENT_HANDLING.tagNameCheck
      // and c) if the attribute name passes whatever the user has configured for CUSTOM_ELEMENT_HANDLING.attributeNameCheck
      _isBasicCustomElement(lcTag) && (CUSTOM_ELEMENT_HANDLING.tagNameCheck instanceof RegExp && regExpTest(CUSTOM_ELEMENT_HANDLING.tagNameCheck, lcTag) || CUSTOM_ELEMENT_HANDLING.tagNameCheck instanceof Function && CUSTOM_ELEMENT_HANDLING.tagNameCheck(lcTag)) && (CUSTOM_ELEMENT_HANDLING.attributeNameCheck instanceof RegExp && regExpTest(CUSTOM_ELEMENT_HANDLING.attributeNameCheck, lcName) || CUSTOM_ELEMENT_HANDLING.attributeNameCheck instanceof Function && CUSTOM_ELEMENT_HANDLING.attributeNameCheck(lcName, lcTag)) ||
      // Alternative, second condition checks if it's an `is`-attribute, AND
      // the value passes whatever the user has configured for CUSTOM_ELEMENT_HANDLING.tagNameCheck
      lcName === 'is' && CUSTOM_ELEMENT_HANDLING.allowCustomizedBuiltInElements && (CUSTOM_ELEMENT_HANDLING.tagNameCheck instanceof RegExp && regExpTest(CUSTOM_ELEMENT_HANDLING.tagNameCheck, value) || CUSTOM_ELEMENT_HANDLING.tagNameCheck instanceof Function && CUSTOM_ELEMENT_HANDLING.tagNameCheck(value))) ; else {
        return false;
      }
      /* Check value is safe. First, is attr inert? If so, is safe */
    } else if (URI_SAFE_ATTRIBUTES[lcName]) ; else if (regExpTest(IS_ALLOWED_URI$1, stringReplace(value, ATTR_WHITESPACE$1, ''))) ; else if ((lcName === 'src' || lcName === 'xlink:href' || lcName === 'href') && lcTag !== 'script' && stringIndexOf(value, 'data:') === 0 && DATA_URI_TAGS[lcTag]) ; else if (ALLOW_UNKNOWN_PROTOCOLS && !regExpTest(IS_SCRIPT_OR_DATA$1, stringReplace(value, ATTR_WHITESPACE$1, ''))) ; else if (value) {
      return false;
    } else ;
    return true;
  };
  /* Names the HTML spec reserves from valid-custom-element-name; these must
   * never be treated as basic custom elements even when a permissive
   * CUSTOM_ELEMENT_HANDLING.tagNameCheck is configured. */
  const RESERVED_CUSTOM_ELEMENT_NAMES = addToSet({}, ['annotation-xml', 'color-profile', 'font-face', 'font-face-format', 'font-face-name', 'font-face-src', 'font-face-uri', 'missing-glyph']);
  /**
   * _isBasicCustomElement
   * checks if at least one dash is included in tagName, and it's not the first char
   * for more sophisticated checking see https://github.com/sindresorhus/validate-element-name
   *
   * @param tagName name of the tag of the node to sanitize
   * @returns Returns true if the tag name meets the basic criteria for a custom element, otherwise false.
   */
  const _isBasicCustomElement = function _isBasicCustomElement(tagName) {
    return !RESERVED_CUSTOM_ELEMENT_NAMES[stringToLowerCase(tagName)] && regExpTest(CUSTOM_ELEMENT$1, tagName);
  };
  /**
   * Wrap an attribute value in the matching Trusted Types object when
   * the active policy requires it. Namespaced attributes pass through
   * unchanged (no TT support yet, see
   * https://bugs.chromium.org/p/chromium/issues/detail?id=1305293).
   *
   * @param lcTag lowercase tag name of the containing element
   * @param lcName lowercase attribute name
   * @param namespaceURI the attribute's namespace, if any
   * @param value the attribute value to wrap
   * @return the value, wrapped when Trusted Types demand it
   */
  const _applyTrustedTypesToAttribute = function _applyTrustedTypesToAttribute(lcTag, lcName, namespaceURI, value) {
    if (trustedTypesPolicy && typeof trustedTypes === 'object' && typeof trustedTypes.getAttributeType === 'function' && !namespaceURI) {
      switch (trustedTypes.getAttributeType(lcTag, lcName)) {
        case 'TrustedHTML':
          {
            return _createTrustedHTML(value);
          }
        case 'TrustedScriptURL':
          {
            return _createTrustedScriptURL(value);
          }
      }
    }
    return value;
  };
  /**
   * Write a modified attribute value back onto the element. On
   * success, re-probe for clobbering introduced by the new value and
   * remove the element when found; otherwise pop the removal entry
   * recorded by the earlier _removeAttribute (long-standing pairing
   * with the SANITIZE_NAMED_PROPS path - do not "fix" casually). On
   * failure, remove the attribute instead.
   *
   * @param currentNode the element carrying the attribute
   * @param name the attribute name as present on the element
   * @param namespaceURI the attribute's namespace, if any
   * @param value the new attribute value
   */
  const _setAttributeValue = function _setAttributeValue(currentNode, name, namespaceURI, value) {
    try {
      if (namespaceURI) {
        currentNode.setAttributeNS(namespaceURI, name, value);
      } else {
        /* Fallback to setAttribute() for browser-unrecognized namespaces e.g. "x-schema". */
        currentNode.setAttribute(name, value);
      }
      if (_isClobbered(currentNode)) {
        _forceRemove(currentNode);
      } else {
        arrayPop(DOMPurify.removed);
      }
    } catch (_) {
      _removeAttribute(name, currentNode);
    }
  };
  /**
   * _sanitizeAttributes
   *
   * @protect attributes
   * @protect nodeName
   * @protect removeAttribute
   * @protect setAttribute
   *
   * @param currentNode to sanitize
   */
  const _sanitizeAttributes = function _sanitizeAttributes(currentNode) {
    /* Execute a hook if present */
    _executeHooks(hooks.beforeSanitizeAttributes, currentNode, null);
    const attributes = currentNode.attributes;
    /* Check if we have attributes; if not we might have a text node */
    if (!attributes || _isClobbered(currentNode)) {
      return;
    }
    /* Same lazy-install guard as uponSanitizeElement (see there): fork the
       attribute allowlist off its shared binding before a hook can see it. */
    ALLOWED_ATTR = _forkSharedAllowlist(hooks.uponSanitizeAttribute, ALLOWED_ATTR, DEFAULT_ALLOWED_ATTR, SET_CONFIG_ALLOWED_ATTR);
    const hookEvent = {
      attrName: '',
      attrValue: '',
      keepAttr: true,
      allowedAttributes: ALLOWED_ATTR,
      forceKeepAttr: undefined
    };
    let l = attributes.length;
    const lcTag = transformCaseFunc(currentNode.nodeName);
    /* Go backwards over all attributes; safely remove bad ones */
    while (l--) {
      const attr = attributes[l];
      const name = attr.name,
        namespaceURI = attr.namespaceURI,
        attrValue = attr.value;
      const lcName = transformCaseFunc(name);
      const initValue = attrValue;
      let value = name === 'value' ? initValue : stringTrim(initValue);
      /* Execute a hook if present */
      hookEvent.attrName = lcName;
      hookEvent.attrValue = value;
      hookEvent.keepAttr = true;
      hookEvent.forceKeepAttr = undefined; // Allows developers to see this is a property they can set
      _executeHooks(hooks.uponSanitizeAttribute, currentNode, hookEvent);
      value = hookEvent.attrValue;
      /* Full DOM Clobbering protection via namespace isolation,
       * Prefix id and name attributes with `user-content-`
       */
      if (SANITIZE_NAMED_PROPS && (lcName === 'id' || lcName === 'name') && stringIndexOf(value, SANITIZE_NAMED_PROPS_PREFIX) !== 0) {
        // Remove the attribute with this value
        _removeAttribute(name, currentNode);
        // Prefix the value and later re-create the attribute with the sanitized value
        value = SANITIZE_NAMED_PROPS_PREFIX + value;
      }
      // Else: already prefixed, leave the attribute alone — the prefix is
      // itself the clobbering protection, and re-applying it is incorrect.
      /* Work around a security issue with comments inside attributes */
      if (SAFE_FOR_XML && regExpTest(/((--!?|])>)|<\/(style|script|title|xmp|textarea|noscript|iframe|noembed|noframes)/i, value)) {
        _removeAttribute(name, currentNode);
        continue;
      }
      /* Make sure we cannot easily use animated hrefs, even if animations are allowed */
      if (lcName === 'attributename' && stringMatch(value, 'href')) {
        _removeAttribute(name, currentNode);
        continue;
      }
      /* Did the hooks force-keep the attribute? */
      if (hookEvent.forceKeepAttr) {
        continue;
      }
      /* Did the hooks approve of the attribute? */
      if (!hookEvent.keepAttr) {
        _removeAttribute(name, currentNode);
        continue;
      }
      /* Work around a security issue in jQuery 3.0 */
      if (!ALLOW_SELF_CLOSE_IN_ATTR && regExpTest(SELF_CLOSING_TAG, value)) {
        _removeAttribute(name, currentNode);
        continue;
      }
      /* Sanitize attribute content to be template-safe */
      if (SAFE_FOR_TEMPLATES) {
        value = _stripTemplateExpressions(value);
      }
      /* Is `value` valid for this attribute? */
      if (!_isValidAttribute(lcTag, lcName, value)) {
        _removeAttribute(name, currentNode);
        continue;
      }
      /* Handle attributes that require Trusted Types */
      value = _applyTrustedTypesToAttribute(lcTag, lcName, namespaceURI, value);
      /* Handle invalid data-* attribute set by try-catching it */
      if (value !== initValue) {
        _setAttributeValue(currentNode, name, namespaceURI, value);
      }
    }
    /* Execute a hook if present */
    _executeHooks(hooks.afterSanitizeAttributes, currentNode, null);
  };
  /**
   * _sanitizeShadowDOM
   *
   * @param fragment to iterate over recursively
   */
  const _sanitizeShadowDOM2 = function _sanitizeShadowDOM(fragment) {
    let shadowNode = null;
    const shadowIterator = _createNodeIterator(fragment);
    /* Execute a hook if present */
    _executeHooks(hooks.beforeSanitizeShadowDOM, fragment, null);
    while (shadowNode = shadowIterator.nextNode()) {
      /* Execute a hook if present */
      _executeHooks(hooks.uponSanitizeShadowNode, shadowNode, null);
      /* Sanitize tags and elements */
      _sanitizeElements(shadowNode, fragment);
      /* Check attributes next */
      _sanitizeAttributes(shadowNode);
      /* Deep shadow DOM detected.
         Realm-safe check (GHSA-hpcv-96wg-7vj8): use nodeType against the
         DOCUMENT_FRAGMENT_NODE constant rather than instanceof, so we
         recurse into <template>.content from foreign realms too. */
      if (_isDocumentFragment(shadowNode.content)) {
        _sanitizeShadowDOM2(shadowNode.content);
      }
      /* An element iterated here may itself host an attached
         shadow root. The default NodeIterator does not enter shadow
         trees, so a shadow root nested inside template.content was
         previously reached by no walk at all (the pre-pass at
         _sanitizeAttachedShadowRoots descends via childNodes, which
         doesn't enter template.content; the template-content recursion
         above iterates the content but never inspected shadowRoot).
         Walk it explicitly. The nodeType guard avoids reading
         shadowRoot off text / comment / CDATA / PI nodes that the
         iterator also surfaces. */
      const shadowNodeType = getNodeType ? getNodeType(shadowNode) : shadowNode.nodeType;
      if (shadowNodeType === NODE_TYPE.element) {
        const innerSr = getShadowRoot(shadowNode);
        if (_isDocumentFragment(innerSr)) {
          _sanitizeAttachedShadowRoots(innerSr);
          _sanitizeShadowDOM2(innerSr);
        }
      }
    }
    /* Execute a hook if present */
    _executeHooks(hooks.afterSanitizeShadowDOM, fragment, null);
  };
  /**
   * _sanitizeAttachedShadowRoots
   *
   * Walks `root` and feeds every attached shadow root we encounter into
   * the existing _sanitizeShadowDOM pipeline. The default node iterator
   * does not descend into shadow trees, so nodes inside an attached
   * shadow root would otherwise be skipped entirely.
   *
   * Two real input paths put attached shadow roots in front of us:
   *   1. IN_PLACE on a DOM node that already has shadow roots attached.
   *   2. DOM-node input where importNode(dirty, true) deep-clones the
   *      shadow root because it was created with `clonable: true`.
   *
   * This pass runs once, up front, so the main iteration loop (and the
   * existing _sanitizeShadowDOM template-content recursion) stay
   * untouched — string-input paths are not affected.
   *
   * @param root the subtree root to walk for attached shadow roots
   */
  const _sanitizeAttachedShadowRoots = function _sanitizeAttachedShadowRoots(root) {
    /* Iterative (explicit stack) rather than per-child recursion. DOM APIs
       impose no depth cap, so an attacker-shaped tree (JSON/CRDT/editor data
       built straight into the DOM — the IN_PLACE surface) deeper than the JS
       call-stack budget would otherwise overflow native recursion here and
       throw at the IN_PLACE entry pre-pass, before a single node is
       sanitized, leaving the caller's live tree untouched (fail-open). See
       campaign-3 F4. A heap stack keeps depth off the call stack.
            Each work item is either a node to descend into, or a deferred
       `_sanitizeShadowDOM` for an already-walked shadow root. The deferred
       form preserves the original post-order discipline: a shadow root's
       nested shadow roots are discovered before the outer shadow is
       sanitized (which may remove hosts). Pushes are in reverse of the
       desired processing order (LIFO): template content, then children, then
       the shadow-sanitize, then the shadow walk — so the order matches the
       previous recursion exactly. */
    const stack = [{
      node: root,
      shadow: null
    }];
    while (stack.length > 0) {
      const item = stack.pop();
      /* Deferred shadow-DOM sanitisation: runs after its subtree was walked. */
      if (item.shadow) {
        _sanitizeShadowDOM2(item.shadow);
        continue;
      }
      const node = item.node;
      const nodeType = getNodeType ? getNodeType(node) : node.nodeType;
      const isElement = nodeType === NODE_TYPE.element;
      /* (pushed last → processed first) Children, snapshotted in reverse so
         the first child is processed first. Snapshotting matters because a
         hook may detach siblings mid-walk. */
      const childNodes = getChildNodes(node);
      if (childNodes) {
        for (let i = childNodes.length - 1; i >= 0; --i) {
          stack.push({
            node: childNodes[i],
            shadow: null
          });
        }
      }
      /* (pushed before children → processed after them, matching the old
         "template content last" order) When the node is a <template>,
         descend into its content. */
      if (isElement) {
        const rootName = getNodeName ? getNodeName(node) : null;
        if (typeof rootName === 'string' && transformCaseFunc(rootName) === 'template') {
          const content = node.content;
          if (_isDocumentFragment(content)) {
            stack.push({
              node: content,
              shadow: null
            });
          }
        }
      }
      /* Shadow root (processed first): walk its subtree, then sanitise it.
         Realm-safe check (GHSA-hpcv-96wg-7vj8): nodeType-based detection
         rather than `instanceof DocumentFragment`, which is realm-bound and
         silently skipped foreign-realm shadow roots (e.g.
         iframe.contentDocument attachShadow). */
      if (isElement) {
        const sr = getShadowRoot(node);
        if (_isDocumentFragment(sr)) {
          /* Push the deferred sanitise first so it pops after the shadow
             walk we push next, i.e. nested shadow roots are discovered
             before this one is sanitised. */
          stack.push({
            node: null,
            shadow: sr
          }, {
            node: sr,
            shadow: null
          });
        }
      }
    }
  };
  // eslint-disable-next-line complexity
  DOMPurify.sanitize = function (dirty) {
    let cfg = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};
    let body = null;
    let importedNode = null;
    let currentNode = null;
    let returnNode = null;
    /* Make sure we have a string to sanitize.
      DO NOT return early, as this will return the wrong type if
      the user has requested a DOM object rather than a string */
    IS_EMPTY_INPUT = !dirty;
    if (IS_EMPTY_INPUT) {
      dirty = '<!-->';
    }
    /* Stringify, in case dirty is an object */
    if (typeof dirty !== 'string' && !_isNode(dirty)) {
      dirty = stringifyValue(dirty);
      if (typeof dirty !== 'string') {
        throw typeErrorCreate('dirty is not a string, aborting');
      }
    }
    /* Return dirty HTML if DOMPurify cannot run */
    if (!DOMPurify.isSupported) {
      return dirty;
    }
    /* Assign config vars */
    if (SET_CONFIG) {
      /* Persistent setConfig() path: _parseConfig is skipped, so the sets are
       * not re-derived per call. Restore them from the pristine bindings
       * captured at setConfig() time so a previous call's hook clone (mutated
       * below) does not carry over. */
      ALLOWED_TAGS = SET_CONFIG_ALLOWED_TAGS;
      ALLOWED_ATTR = SET_CONFIG_ALLOWED_ATTR;
    } else {
      _parseConfig(cfg);
    }
    /* Clone the hook-mutable allowlists before the walk whenever an
     * uponSanitize* hook is registered. The hook event exposes ALLOWED_TAGS
     * and ALLOWED_ATTR by reference (as allowedTags / allowedAttributes), so
     * a hook that widens them would otherwise mutate the shared set
     * permanently: across later calls and across every element. Cloning per
     * walk keeps documented in-call widening working while scoping it to the
     * call. A single guard for both config paths - the per-call path rebinds
     * the sets in _parseConfig each call, the persistent path restores them
     * from the captured bindings just above - so the two cannot diverge. */
    if (hooks.uponSanitizeElement.length > 0 || hooks.uponSanitizeAttribute.length > 0) {
      ALLOWED_TAGS = clone(ALLOWED_TAGS);
    }
    if (hooks.uponSanitizeAttribute.length > 0) {
      ALLOWED_ATTR = clone(ALLOWED_ATTR);
    }
    /* Clean up removed elements */
    DOMPurify.removed = [];
    /* Resolve IN_PLACE for this call without mutating persistent config.
       Writing the IN_PLACE closure variable here leaks under setConfig(),
       where _parseConfig is skipped on later calls: a single string call would
       disable in-place mode for every subsequent node call, returning a
       sanitized copy while leaving the caller's node — which in-place callers
       keep using and whose return value they ignore — unsanitized. REPORT-2. */
    const inPlace = IN_PLACE && typeof dirty !== 'string' && _isNode(dirty);
    if (inPlace) {
      /* Declarative-partial-updates / streaming pre-pass: sever every patch
         linkage across the live tree BEFORE the walk, so no patch can fire
         mid-walk and inject into an already-processed region. Runs first, so
         it also covers the forbidden/clobbered roots that throw below. */
      _neutralizePatchLinkage(dirty);
      /* Do some early pre-sanitization to avoid unsafe root nodes.
         Read nodeName through the cached prototype getter — a clobbering
         child named "nodeName" on the form root would otherwise shadow
         the property and let this check skip the root-allowlist
         validation entirely. */
      const nn = getNodeName ? getNodeName(dirty) : dirty.nodeName;
      if (typeof nn === 'string') {
        const tagName = transformCaseFunc(nn);
        if (!ALLOWED_TAGS[tagName] || FORBID_TAGS[tagName]) {
          /* Fail closed on a live root: neutralize handlers/children before
             throwing, exactly as the mid-walk abort path does. */
          _neutralizeRoot(dirty);
          throw typeErrorCreate('root node is forbidden and cannot be sanitized in-place');
        }
      }
      /* Pre-flight the root through _isClobbered. The iterator-driven
         removal path can not detach a parent-less root: _forceRemove
         falls through to Element.prototype.remove(), which per spec
         is a no-op on a node with no parent. A clobbered root would
         then survive the main loop with its attributes uninspected,
         because _sanitizeAttributes early-returns on _isClobbered. The
         result would be an attacker-controlled form, complete with any
         event-handler attributes the caller passed in, handed back to
         the application unsanitized. Refuse to sanitize such a root
         the same way we refuse a forbidden tag. GHSA-r47g-fvhr-h676. */
      if (_isClobbered(dirty)) {
        /* Fail closed on a live clobbered root before throwing.
           _neutralizeRoot's reads are clobber-safe (cached getters); the
           form's non-clobbered descendants, e.g. an armed <img>, are scrubbed. */
        _neutralizeRoot(dirty);
        throw typeErrorCreate('root node is clobbered and cannot be sanitized in-place');
      }
      /* Sanitize attached shadow roots before the main iterator runs.
         The iterator does not descend into shadow trees. Same fail-closed
         barrier as the main walk (campaign-3 F2): a custom-element reaction
         inside a shadow root could abort this pre-pass before the walk runs,
         which would otherwise leave the entire live tree unsanitized. */
      try {
        _sanitizeAttachedShadowRoots(dirty);
      } catch (error) {
        _neutralizeRoot(dirty);
        throw error;
      }
    } else if (_isNode(dirty)) {
      /* If dirty is a DOM element, append to an empty document to avoid
         elements being stripped by the parser */
      body = _initDocument('<!---->');
      importedNode = body.ownerDocument.importNode(dirty, true);
      if (importedNode.nodeType === NODE_TYPE.element && importedNode.nodeName === 'BODY') {
        /* Node is already a body, use as is */
        body = importedNode;
      } else if (importedNode.nodeName === 'HTML') {
        body = importedNode;
      } else {
        // eslint-disable-next-line unicorn/prefer-dom-node-append
        body.appendChild(importedNode);
      }
      /* Clonable shadow roots are deep-cloned by importNode(); sanitize
         them before the main iterator runs, since the iterator does not
         descend into shadow trees. The walk routes every read through a
         cached prototype getter so clobbering descendants on a form root
         cannot hide a shadow host from this pass. */
      _sanitizeAttachedShadowRoots(importedNode);
    } else {
      /* Exit directly if we have nothing to do */
      if (!RETURN_DOM && !SAFE_FOR_TEMPLATES && !WHOLE_DOCUMENT &&
      // eslint-disable-next-line unicorn/prefer-includes
      dirty.indexOf('<') === -1) {
        return trustedTypesPolicy && RETURN_TRUSTED_TYPE ? _createTrustedHTML(dirty) : dirty;
      }
      /* Initialize the document to work on */
      body = _initDocument(dirty);
      /* Check we have a DOM node from the data */
      if (!body) {
        return RETURN_DOM ? null : RETURN_TRUSTED_TYPE ? emptyHTML : '';
      }
    }
    /* Remove first element node (ours) if FORCE_BODY is set */
    if (body && FORCE_BODY) {
      _forceRemove(body.firstChild);
    }
    /* Get node iterator */
    const walkRoot = inPlace ? dirty : body;
    /* Now start iterating over the created document.
       The walk runs inside an exception barrier (campaign-3 F2): a re-entrant
       engine/custom-element mutation can detach a node mid-walk so
       `_forceRemove`'s parentless guard throws, aborting the loop. Without the
       barrier the caller's in-place tree would be left half-sanitized with the
       unvisited tail still armed. _createNodeIterator itself is inside the
       barrier too: constructing the iterator dereferences the root's document,
       and any failure there (e.g. an exotic/clobbered root) must still fail
       closed rather than skip the neutralize. On any throw we fail closed -
       strip the in-place root bare - then rethrow so the existing throw
       contract is preserved. (String/DOM-copy paths never return the partial
       body, so the propagating throw is already fail-closed there.) */
    try {
      const nodeIterator = _createNodeIterator(walkRoot);
      while (currentNode = nodeIterator.nextNode()) {
        /* Sanitize tags and elements */
        _sanitizeElements(currentNode, walkRoot);
        /* Check attributes next */
        _sanitizeAttributes(currentNode);
        /* Shadow DOM detected, sanitize it.
           Realm-safe check (GHSA-hpcv-96wg-7vj8): nodeType-based detection
           instead of instanceof, so foreign-realm <template>.content is
           walked correctly. */
        if (_isDocumentFragment(currentNode.content)) {
          _sanitizeShadowDOM2(currentNode.content);
        }
      }
    } catch (error) {
      if (inPlace) {
        _neutralizeRoot(dirty);
        /* Nodes _forceRemove'd earlier in the aborted walk are already
           detached from the root, so _neutralizeRoot's subtree pass does not
           reach them. Defuse them too, mirroring the success-path loop below. */
        arrayForEach(DOMPurify.removed, entry => {
          if (entry.element) {
            _neutralizeSubtree(entry.element);
          }
        });
      }
      throw error;
    }
    /* If we sanitized `dirty` in-place, return it. */
    if (inPlace) {
      /* Fail-closed completion of the audit-5 F1 fix: every node removed from
         the caller's live tree is detached but may still hold a queued
         resource-event handler that fires in page scope after we return. The
         move-hoist covers only disallowed-tag KEEP_CONTENT removals; strip the
         non-allow-listed attributes off every other removed subtree (clobber,
         mXSS, namespace, comments, KEEP_CONTENT:false, …) so those handlers are
         cancelled before any event can fire. Runs synchronously, pre-return. */
      arrayForEach(DOMPurify.removed, entry => {
        if (entry.element) {
          _neutralizeSubtree(entry.element);
        }
      });
      if (SAFE_FOR_TEMPLATES) {
        _scrubTemplateExpressions2(dirty);
      }
      return dirty;
    }
    /* Return sanitized string or DOM */
    if (RETURN_DOM) {
      if (SAFE_FOR_TEMPLATES) {
        _scrubTemplateExpressions2(body);
      }
      if (RETURN_DOM_FRAGMENT) {
        returnNode = createDocumentFragment.call(body.ownerDocument);
        while (body.firstChild) {
          // eslint-disable-next-line unicorn/prefer-dom-node-append
          returnNode.appendChild(body.firstChild);
        }
      } else {
        returnNode = body;
      }
      if (ALLOWED_ATTR.shadowroot || ALLOWED_ATTR.shadowrootmode) {
        /*
          AdoptNode() is not used because internal state is not reset
          (e.g. the past names map of a HTMLFormElement), this is safe
          in theory but we would rather not risk another attack vector.
          The state that is cloned by importNode() is explicitly defined
          by the specs.
        */
        returnNode = importNode.call(originalDocument, returnNode, true);
      }
      return returnNode;
    }
    let serializedHTML = WHOLE_DOCUMENT ? body.outerHTML : body.innerHTML;
    /* Serialize doctype if allowed */
    if (WHOLE_DOCUMENT && ALLOWED_TAGS['!doctype'] && body.ownerDocument && body.ownerDocument.doctype && body.ownerDocument.doctype.name && regExpTest(DOCTYPE_NAME, body.ownerDocument.doctype.name)) {
      serializedHTML = '<!DOCTYPE ' + body.ownerDocument.doctype.name + '>\n' + serializedHTML;
    }
    /* Sanitize final string template-safe */
    if (SAFE_FOR_TEMPLATES) {
      serializedHTML = _stripTemplateExpressions(serializedHTML);
    }
    return trustedTypesPolicy && RETURN_TRUSTED_TYPE ? _createTrustedHTML(serializedHTML) : serializedHTML;
  };
  DOMPurify.setConfig = function () {
    let cfg = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
    _parseConfig(cfg);
    SET_CONFIG = true;
    SET_CONFIG_ALLOWED_TAGS = ALLOWED_TAGS;
    SET_CONFIG_ALLOWED_ATTR = ALLOWED_ATTR;
  };
  DOMPurify.clearConfig = function () {
    CONFIG = null;
    SET_CONFIG = false;
    SET_CONFIG_ALLOWED_TAGS = null;
    SET_CONFIG_ALLOWED_ATTR = null;
    // Drop any caller-supplied Trusted Types policy so it cannot poison later
    // `RETURN_TRUSTED_TYPE` output. The internal default policy (cached, and
    // never recreated — Trusted Types throws on duplicate names) is restored by
    // the next `_parseConfig`. See GHSA-vxr8-fq34-vvx9.
    trustedTypesPolicy = defaultTrustedTypesPolicy;
    emptyHTML = '';
  };
  DOMPurify.isValidAttribute = function (tag, attr, value) {
    /* Initialize shared config vars if necessary. */
    if (!CONFIG) {
      _parseConfig({});
    }
    const lcTag = transformCaseFunc(tag);
    const lcName = transformCaseFunc(attr);
    return _isValidAttribute(lcTag, lcName, value);
  };
  DOMPurify.addHook = function (entryPoint, hookFunction) {
    if (typeof hookFunction !== 'function') {
      return;
    }
    /* Reject unknown entry points. Without this, a non-hook key (e.g.
     * '__proto__') indexes off the prototype chain rather than a real
     * hook array, and arrayPush then writes to Object.prototype. Guard
     * with an own-property check against the known hook names. */
    if (!objectHasOwnProperty(hooks, entryPoint)) {
      return;
    }
    arrayPush(hooks[entryPoint], hookFunction);
  };
  DOMPurify.removeHook = function (entryPoint, hookFunction) {
    if (!objectHasOwnProperty(hooks, entryPoint)) {
      return undefined;
    }
    if (hookFunction !== undefined) {
      const index = arrayLastIndexOf(hooks[entryPoint], hookFunction);
      return index === -1 ? undefined : arraySplice(hooks[entryPoint], index, 1)[0];
    }
    return arrayPop(hooks[entryPoint]);
  };
  DOMPurify.removeHooks = function (entryPoint) {
    if (!objectHasOwnProperty(hooks, entryPoint)) {
      return;
    }
    hooks[entryPoint] = [];
  };
  DOMPurify.removeAllHooks = function () {
    hooks = _createHooksMap();
  };
  return DOMPurify;
}
var purify = createDOMPurify();


//# sourceMappingURL=purify.es.mjs.map

;// ./node_modules/marked/lib/marked.esm.js
/**
 * marked v15.0.12 - a markdown parser
 * Copyright (c) 2011-2025, Christopher Jeffrey. (MIT Licensed)
 * https://github.com/markedjs/marked
 */

/**
 * DO NOT EDIT THIS FILE
 * The code in this file is generated from files in ./src/
 */


// src/defaults.ts
function _getDefaults() {
  return {
    async: false,
    breaks: false,
    extensions: null,
    gfm: true,
    hooks: null,
    pedantic: false,
    renderer: null,
    silent: false,
    tokenizer: null,
    walkTokens: null
  };
}
var _defaults = _getDefaults();
function changeDefaults(newDefaults) {
  _defaults = newDefaults;
}

// src/rules.ts
var noopTest = { exec: () => null };
function edit(regex, opt = "") {
  let source = typeof regex === "string" ? regex : regex.source;
  const obj = {
    replace: (name, val) => {
      let valSource = typeof val === "string" ? val : val.source;
      valSource = valSource.replace(other.caret, "$1");
      source = source.replace(name, valSource);
      return obj;
    },
    getRegex: () => {
      return new RegExp(source, opt);
    }
  };
  return obj;
}
var other = {
  codeRemoveIndent: /^(?: {1,4}| {0,3}\t)/gm,
  outputLinkReplace: /\\([\[\]])/g,
  indentCodeCompensation: /^(\s+)(?:```)/,
  beginningSpace: /^\s+/,
  endingHash: /#$/,
  startingSpaceChar: /^ /,
  endingSpaceChar: / $/,
  nonSpaceChar: /[^ ]/,
  newLineCharGlobal: /\n/g,
  tabCharGlobal: /\t/g,
  multipleSpaceGlobal: /\s+/g,
  blankLine: /^[ \t]*$/,
  doubleBlankLine: /\n[ \t]*\n[ \t]*$/,
  blockquoteStart: /^ {0,3}>/,
  blockquoteSetextReplace: /\n {0,3}((?:=+|-+) *)(?=\n|$)/g,
  blockquoteSetextReplace2: /^ {0,3}>[ \t]?/gm,
  listReplaceTabs: /^\t+/,
  listReplaceNesting: /^ {1,4}(?=( {4})*[^ ])/g,
  listIsTask: /^\[[ xX]\] /,
  listReplaceTask: /^\[[ xX]\] +/,
  anyLine: /\n.*\n/,
  hrefBrackets: /^<(.*)>$/,
  tableDelimiter: /[:|]/,
  tableAlignChars: /^\||\| *$/g,
  tableRowBlankLine: /\n[ \t]*$/,
  tableAlignRight: /^ *-+: *$/,
  tableAlignCenter: /^ *:-+: *$/,
  tableAlignLeft: /^ *:-+ *$/,
  startATag: /^<a /i,
  endATag: /^<\/a>/i,
  startPreScriptTag: /^<(pre|code|kbd|script)(\s|>)/i,
  endPreScriptTag: /^<\/(pre|code|kbd|script)(\s|>)/i,
  startAngleBracket: /^</,
  endAngleBracket: />$/,
  pedanticHrefTitle: /^([^'"]*[^\s])\s+(['"])(.*)\2/,
  unicodeAlphaNumeric: /[\p{L}\p{N}]/u,
  escapeTest: /[&<>"']/,
  escapeReplace: /[&<>"']/g,
  escapeTestNoEncode: /[<>"']|&(?!(#\d{1,7}|#[Xx][a-fA-F0-9]{1,6}|\w+);)/,
  escapeReplaceNoEncode: /[<>"']|&(?!(#\d{1,7}|#[Xx][a-fA-F0-9]{1,6}|\w+);)/g,
  unescapeTest: /&(#(?:\d+)|(?:#x[0-9A-Fa-f]+)|(?:\w+));?/ig,
  caret: /(^|[^\[])\^/g,
  percentDecode: /%25/g,
  findPipe: /\|/g,
  splitPipe: / \|/,
  slashPipe: /\\\|/g,
  carriageReturn: /\r\n|\r/g,
  spaceLine: /^ +$/gm,
  notSpaceStart: /^\S*/,
  endingNewline: /\n$/,
  listItemRegex: (bull) => new RegExp(`^( {0,3}${bull})((?:[	 ][^\\n]*)?(?:\\n|$))`),
  nextBulletRegex: (indent) => new RegExp(`^ {0,${Math.min(3, indent - 1)}}(?:[*+-]|\\d{1,9}[.)])((?:[ 	][^\\n]*)?(?:\\n|$))`),
  hrRegex: (indent) => new RegExp(`^ {0,${Math.min(3, indent - 1)}}((?:- *){3,}|(?:_ *){3,}|(?:\\* *){3,})(?:\\n+|$)`),
  fencesBeginRegex: (indent) => new RegExp(`^ {0,${Math.min(3, indent - 1)}}(?:\`\`\`|~~~)`),
  headingBeginRegex: (indent) => new RegExp(`^ {0,${Math.min(3, indent - 1)}}#`),
  htmlBeginRegex: (indent) => new RegExp(`^ {0,${Math.min(3, indent - 1)}}<(?:[a-z].*>|!--)`, "i")
};
var newline = /^(?:[ \t]*(?:\n|$))+/;
var blockCode = /^((?: {4}| {0,3}\t)[^\n]+(?:\n(?:[ \t]*(?:\n|$))*)?)+/;
var fences = /^ {0,3}(`{3,}(?=[^`\n]*(?:\n|$))|~{3,})([^\n]*)(?:\n|$)(?:|([\s\S]*?)(?:\n|$))(?: {0,3}\1[~`]* *(?=\n|$)|$)/;
var hr = /^ {0,3}((?:-[\t ]*){3,}|(?:_[ \t]*){3,}|(?:\*[ \t]*){3,})(?:\n+|$)/;
var heading = /^ {0,3}(#{1,6})(?=\s|$)(.*)(?:\n+|$)/;
var bullet = /(?:[*+-]|\d{1,9}[.)])/;
var lheadingCore = /^(?!bull |blockCode|fences|blockquote|heading|html|table)((?:.|\n(?!\s*?\n|bull |blockCode|fences|blockquote|heading|html|table))+?)\n {0,3}(=+|-+) *(?:\n+|$)/;
var lheading = edit(lheadingCore).replace(/bull/g, bullet).replace(/blockCode/g, /(?: {4}| {0,3}\t)/).replace(/fences/g, / {0,3}(?:`{3,}|~{3,})/).replace(/blockquote/g, / {0,3}>/).replace(/heading/g, / {0,3}#{1,6}/).replace(/html/g, / {0,3}<[^\n>]+>\n/).replace(/\|table/g, "").getRegex();
var lheadingGfm = edit(lheadingCore).replace(/bull/g, bullet).replace(/blockCode/g, /(?: {4}| {0,3}\t)/).replace(/fences/g, / {0,3}(?:`{3,}|~{3,})/).replace(/blockquote/g, / {0,3}>/).replace(/heading/g, / {0,3}#{1,6}/).replace(/html/g, / {0,3}<[^\n>]+>\n/).replace(/table/g, / {0,3}\|?(?:[:\- ]*\|)+[\:\- ]*\n/).getRegex();
var _paragraph = /^([^\n]+(?:\n(?!hr|heading|lheading|blockquote|fences|list|html|table| +\n)[^\n]+)*)/;
var blockText = /^[^\n]+/;
var _blockLabel = /(?!\s*\])(?:\\.|[^\[\]\\])+/;
var def = edit(/^ {0,3}\[(label)\]: *(?:\n[ \t]*)?([^<\s][^\s]*|<.*?>)(?:(?: +(?:\n[ \t]*)?| *\n[ \t]*)(title))? *(?:\n+|$)/).replace("label", _blockLabel).replace("title", /(?:"(?:\\"?|[^"\\])*"|'[^'\n]*(?:\n[^'\n]+)*\n?'|\([^()]*\))/).getRegex();
var list = edit(/^( {0,3}bull)([ \t][^\n]+?)?(?:\n|$)/).replace(/bull/g, bullet).getRegex();
var _tag = "address|article|aside|base|basefont|blockquote|body|caption|center|col|colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|frame|frameset|h[1-6]|head|header|hr|html|iframe|legend|li|link|main|menu|menuitem|meta|nav|noframes|ol|optgroup|option|p|param|search|section|summary|table|tbody|td|tfoot|th|thead|title|tr|track|ul";
var _comment = /<!--(?:-?>|[\s\S]*?(?:-->|$))/;
var marked_esm_html = edit(
  "^ {0,3}(?:<(script|pre|style|textarea)[\\s>][\\s\\S]*?(?:</\\1>[^\\n]*\\n+|$)|comment[^\\n]*(\\n+|$)|<\\?[\\s\\S]*?(?:\\?>\\n*|$)|<![A-Z][\\s\\S]*?(?:>\\n*|$)|<!\\[CDATA\\[[\\s\\S]*?(?:\\]\\]>\\n*|$)|</?(tag)(?: +|\\n|/?>)[\\s\\S]*?(?:(?:\\n[ 	]*)+\\n|$)|<(?!script|pre|style|textarea)([a-z][\\w-]*)(?:attribute)*? */?>(?=[ \\t]*(?:\\n|$))[\\s\\S]*?(?:(?:\\n[ 	]*)+\\n|$)|</(?!script|pre|style|textarea)[a-z][\\w-]*\\s*>(?=[ \\t]*(?:\\n|$))[\\s\\S]*?(?:(?:\\n[ 	]*)+\\n|$))",
  "i"
).replace("comment", _comment).replace("tag", _tag).replace("attribute", / +[a-zA-Z:_][\w.:-]*(?: *= *"[^"\n]*"| *= *'[^'\n]*'| *= *[^\s"'=<>`]+)?/).getRegex();
var paragraph = edit(_paragraph).replace("hr", hr).replace("heading", " {0,3}#{1,6}(?:\\s|$)").replace("|lheading", "").replace("|table", "").replace("blockquote", " {0,3}>").replace("fences", " {0,3}(?:`{3,}(?=[^`\\n]*\\n)|~{3,})[^\\n]*\\n").replace("list", " {0,3}(?:[*+-]|1[.)]) ").replace("html", "</?(?:tag)(?: +|\\n|/?>)|<(?:script|pre|style|textarea|!--)").replace("tag", _tag).getRegex();
var blockquote = edit(/^( {0,3}> ?(paragraph|[^\n]*)(?:\n|$))+/).replace("paragraph", paragraph).getRegex();
var blockNormal = {
  blockquote,
  code: blockCode,
  def,
  fences,
  heading,
  hr,
  html: marked_esm_html,
  lheading,
  list,
  newline,
  paragraph,
  table: noopTest,
  text: blockText
};
var gfmTable = edit(
  "^ *([^\\n ].*)\\n {0,3}((?:\\| *)?:?-+:? *(?:\\| *:?-+:? *)*(?:\\| *)?)(?:\\n((?:(?! *\\n|hr|heading|blockquote|code|fences|list|html).*(?:\\n|$))*)\\n*|$)"
).replace("hr", hr).replace("heading", " {0,3}#{1,6}(?:\\s|$)").replace("blockquote", " {0,3}>").replace("code", "(?: {4}| {0,3}	)[^\\n]").replace("fences", " {0,3}(?:`{3,}(?=[^`\\n]*\\n)|~{3,})[^\\n]*\\n").replace("list", " {0,3}(?:[*+-]|1[.)]) ").replace("html", "</?(?:tag)(?: +|\\n|/?>)|<(?:script|pre|style|textarea|!--)").replace("tag", _tag).getRegex();
var blockGfm = {
  ...blockNormal,
  lheading: lheadingGfm,
  table: gfmTable,
  paragraph: edit(_paragraph).replace("hr", hr).replace("heading", " {0,3}#{1,6}(?:\\s|$)").replace("|lheading", "").replace("table", gfmTable).replace("blockquote", " {0,3}>").replace("fences", " {0,3}(?:`{3,}(?=[^`\\n]*\\n)|~{3,})[^\\n]*\\n").replace("list", " {0,3}(?:[*+-]|1[.)]) ").replace("html", "</?(?:tag)(?: +|\\n|/?>)|<(?:script|pre|style|textarea|!--)").replace("tag", _tag).getRegex()
};
var blockPedantic = {
  ...blockNormal,
  html: edit(
    `^ *(?:comment *(?:\\n|\\s*$)|<(tag)[\\s\\S]+?</\\1> *(?:\\n{2,}|\\s*$)|<tag(?:"[^"]*"|'[^']*'|\\s[^'"/>\\s]*)*?/?> *(?:\\n{2,}|\\s*$))`
  ).replace("comment", _comment).replace(/tag/g, "(?!(?:a|em|strong|small|s|cite|q|dfn|abbr|data|time|code|var|samp|kbd|sub|sup|i|b|u|mark|ruby|rt|rp|bdi|bdo|span|br|wbr|ins|del|img)\\b)\\w+(?!:|[^\\w\\s@]*@)\\b").getRegex(),
  def: /^ *\[([^\]]+)\]: *<?([^\s>]+)>?(?: +(["(][^\n]+[")]))? *(?:\n+|$)/,
  heading: /^(#{1,6})(.*)(?:\n+|$)/,
  fences: noopTest,
  // fences not supported
  lheading: /^(.+?)\n {0,3}(=+|-+) *(?:\n+|$)/,
  paragraph: edit(_paragraph).replace("hr", hr).replace("heading", " *#{1,6} *[^\n]").replace("lheading", lheading).replace("|table", "").replace("blockquote", " {0,3}>").replace("|fences", "").replace("|list", "").replace("|html", "").replace("|tag", "").getRegex()
};
var marked_esm_escape = /^\\([!"#$%&'()*+,\-./:;<=>?@\[\]\\^_`{|}~])/;
var inlineCode = /^(`+)([^`]|[^`][\s\S]*?[^`])\1(?!`)/;
var br = /^( {2,}|\\)\n(?!\s*$)/;
var inlineText = /^(`+|[^`])(?:(?= {2,}\n)|[\s\S]*?(?:(?=[\\<!\[`*_]|\b_|$)|[^ ](?= {2,}\n)))/;
var _punctuation = /[\p{P}\p{S}]/u;
var _punctuationOrSpace = /[\s\p{P}\p{S}]/u;
var _notPunctuationOrSpace = /[^\s\p{P}\p{S}]/u;
var punctuation = edit(/^((?![*_])punctSpace)/, "u").replace(/punctSpace/g, _punctuationOrSpace).getRegex();
var _punctuationGfmStrongEm = /(?!~)[\p{P}\p{S}]/u;
var _punctuationOrSpaceGfmStrongEm = /(?!~)[\s\p{P}\p{S}]/u;
var _notPunctuationOrSpaceGfmStrongEm = /(?:[^\s\p{P}\p{S}]|~)/u;
var blockSkip = /\[[^[\]]*?\]\((?:\\.|[^\\\(\)]|\((?:\\.|[^\\\(\)])*\))*\)|`[^`]*?`|<[^<>]*?>/g;
var emStrongLDelimCore = /^(?:\*+(?:((?!\*)punct)|[^\s*]))|^_+(?:((?!_)punct)|([^\s_]))/;
var emStrongLDelim = edit(emStrongLDelimCore, "u").replace(/punct/g, _punctuation).getRegex();
var emStrongLDelimGfm = edit(emStrongLDelimCore, "u").replace(/punct/g, _punctuationGfmStrongEm).getRegex();
var emStrongRDelimAstCore = "^[^_*]*?__[^_*]*?\\*[^_*]*?(?=__)|[^*]+(?=[^*])|(?!\\*)punct(\\*+)(?=[\\s]|$)|notPunctSpace(\\*+)(?!\\*)(?=punctSpace|$)|(?!\\*)punctSpace(\\*+)(?=notPunctSpace)|[\\s](\\*+)(?!\\*)(?=punct)|(?!\\*)punct(\\*+)(?!\\*)(?=punct)|notPunctSpace(\\*+)(?=notPunctSpace)";
var emStrongRDelimAst = edit(emStrongRDelimAstCore, "gu").replace(/notPunctSpace/g, _notPunctuationOrSpace).replace(/punctSpace/g, _punctuationOrSpace).replace(/punct/g, _punctuation).getRegex();
var emStrongRDelimAstGfm = edit(emStrongRDelimAstCore, "gu").replace(/notPunctSpace/g, _notPunctuationOrSpaceGfmStrongEm).replace(/punctSpace/g, _punctuationOrSpaceGfmStrongEm).replace(/punct/g, _punctuationGfmStrongEm).getRegex();
var emStrongRDelimUnd = edit(
  "^[^_*]*?\\*\\*[^_*]*?_[^_*]*?(?=\\*\\*)|[^_]+(?=[^_])|(?!_)punct(_+)(?=[\\s]|$)|notPunctSpace(_+)(?!_)(?=punctSpace|$)|(?!_)punctSpace(_+)(?=notPunctSpace)|[\\s](_+)(?!_)(?=punct)|(?!_)punct(_+)(?!_)(?=punct)",
  "gu"
).replace(/notPunctSpace/g, _notPunctuationOrSpace).replace(/punctSpace/g, _punctuationOrSpace).replace(/punct/g, _punctuation).getRegex();
var anyPunctuation = edit(/\\(punct)/, "gu").replace(/punct/g, _punctuation).getRegex();
var autolink = edit(/^<(scheme:[^\s\x00-\x1f<>]*|email)>/).replace("scheme", /[a-zA-Z][a-zA-Z0-9+.-]{1,31}/).replace("email", /[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+(@)[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)+(?![-_])/).getRegex();
var _inlineComment = edit(_comment).replace("(?:-->|$)", "-->").getRegex();
var tag = edit(
  "^comment|^</[a-zA-Z][\\w:-]*\\s*>|^<[a-zA-Z][\\w-]*(?:attribute)*?\\s*/?>|^<\\?[\\s\\S]*?\\?>|^<![a-zA-Z]+\\s[\\s\\S]*?>|^<!\\[CDATA\\[[\\s\\S]*?\\]\\]>"
).replace("comment", _inlineComment).replace("attribute", /\s+[a-zA-Z:_][\w.:-]*(?:\s*=\s*"[^"]*"|\s*=\s*'[^']*'|\s*=\s*[^\s"'=<>`]+)?/).getRegex();
var _inlineLabel = /(?:\[(?:\\.|[^\[\]\\])*\]|\\.|`[^`]*`|[^\[\]\\`])*?/;
var marked_esm_link = edit(/^!?\[(label)\]\(\s*(href)(?:(?:[ \t]*(?:\n[ \t]*)?)(title))?\s*\)/).replace("label", _inlineLabel).replace("href", /<(?:\\.|[^\n<>\\])+>|[^ \t\n\x00-\x1f]*/).replace("title", /"(?:\\"?|[^"\\])*"|'(?:\\'?|[^'\\])*'|\((?:\\\)?|[^)\\])*\)/).getRegex();
var reflink = edit(/^!?\[(label)\]\[(ref)\]/).replace("label", _inlineLabel).replace("ref", _blockLabel).getRegex();
var nolink = edit(/^!?\[(ref)\](?:\[\])?/).replace("ref", _blockLabel).getRegex();
var reflinkSearch = edit("reflink|nolink(?!\\()", "g").replace("reflink", reflink).replace("nolink", nolink).getRegex();
var inlineNormal = {
  _backpedal: noopTest,
  // only used for GFM url
  anyPunctuation,
  autolink,
  blockSkip,
  br,
  code: inlineCode,
  del: noopTest,
  emStrongLDelim,
  emStrongRDelimAst,
  emStrongRDelimUnd,
  escape: marked_esm_escape,
  link: marked_esm_link,
  nolink,
  punctuation,
  reflink,
  reflinkSearch,
  tag,
  text: inlineText,
  url: noopTest
};
var inlinePedantic = {
  ...inlineNormal,
  link: edit(/^!?\[(label)\]\((.*?)\)/).replace("label", _inlineLabel).getRegex(),
  reflink: edit(/^!?\[(label)\]\s*\[([^\]]*)\]/).replace("label", _inlineLabel).getRegex()
};
var inlineGfm = {
  ...inlineNormal,
  emStrongRDelimAst: emStrongRDelimAstGfm,
  emStrongLDelim: emStrongLDelimGfm,
  url: edit(/^((?:ftp|https?):\/\/|www\.)(?:[a-zA-Z0-9\-]+\.?)+[^\s<]*|^email/, "i").replace("email", /[A-Za-z0-9._+-]+(@)[a-zA-Z0-9-_]+(?:\.[a-zA-Z0-9-_]*[a-zA-Z0-9])+(?![-_])/).getRegex(),
  _backpedal: /(?:[^?!.,:;*_'"~()&]+|\([^)]*\)|&(?![a-zA-Z0-9]+;$)|[?!.,:;*_'"~)]+(?!$))+/,
  del: /^(~~?)(?=[^\s~])((?:\\.|[^\\])*?(?:\\.|[^\s~\\]))\1(?=[^~]|$)/,
  text: /^([`~]+|[^`~])(?:(?= {2,}\n)|(?=[a-zA-Z0-9.!#$%&'*+\/=?_`{\|}~-]+@)|[\s\S]*?(?:(?=[\\<!\[`*~_]|\b_|https?:\/\/|ftp:\/\/|www\.|$)|[^ ](?= {2,}\n)|[^a-zA-Z0-9.!#$%&'*+\/=?_`{\|}~-](?=[a-zA-Z0-9.!#$%&'*+\/=?_`{\|}~-]+@)))/
};
var inlineBreaks = {
  ...inlineGfm,
  br: edit(br).replace("{2,}", "*").getRegex(),
  text: edit(inlineGfm.text).replace("\\b_", "\\b_| {2,}\\n").replace(/\{2,\}/g, "*").getRegex()
};
var block = {
  normal: blockNormal,
  gfm: blockGfm,
  pedantic: blockPedantic
};
var inline = {
  normal: inlineNormal,
  gfm: inlineGfm,
  breaks: inlineBreaks,
  pedantic: inlinePedantic
};

// src/helpers.ts
var escapeReplacements = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&#39;"
};
var getEscapeReplacement = (ch) => escapeReplacements[ch];
function escape2(html2, encode) {
  if (encode) {
    if (other.escapeTest.test(html2)) {
      return html2.replace(other.escapeReplace, getEscapeReplacement);
    }
  } else {
    if (other.escapeTestNoEncode.test(html2)) {
      return html2.replace(other.escapeReplaceNoEncode, getEscapeReplacement);
    }
  }
  return html2;
}
function cleanUrl(href) {
  try {
    href = encodeURI(href).replace(other.percentDecode, "%");
  } catch {
    return null;
  }
  return href;
}
function splitCells(tableRow, count) {
  const row = tableRow.replace(other.findPipe, (match, offset, str) => {
    let escaped = false;
    let curr = offset;
    while (--curr >= 0 && str[curr] === "\\") escaped = !escaped;
    if (escaped) {
      return "|";
    } else {
      return " |";
    }
  }), cells = row.split(other.splitPipe);
  let i = 0;
  if (!cells[0].trim()) {
    cells.shift();
  }
  if (cells.length > 0 && !cells.at(-1)?.trim()) {
    cells.pop();
  }
  if (count) {
    if (cells.length > count) {
      cells.splice(count);
    } else {
      while (cells.length < count) cells.push("");
    }
  }
  for (; i < cells.length; i++) {
    cells[i] = cells[i].trim().replace(other.slashPipe, "|");
  }
  return cells;
}
function rtrim(str, c, invert) {
  const l = str.length;
  if (l === 0) {
    return "";
  }
  let suffLen = 0;
  while (suffLen < l) {
    const currChar = str.charAt(l - suffLen - 1);
    if (currChar === c && !invert) {
      suffLen++;
    } else if (currChar !== c && invert) {
      suffLen++;
    } else {
      break;
    }
  }
  return str.slice(0, l - suffLen);
}
function findClosingBracket(str, b) {
  if (str.indexOf(b[1]) === -1) {
    return -1;
  }
  let level = 0;
  for (let i = 0; i < str.length; i++) {
    if (str[i] === "\\") {
      i++;
    } else if (str[i] === b[0]) {
      level++;
    } else if (str[i] === b[1]) {
      level--;
      if (level < 0) {
        return i;
      }
    }
  }
  if (level > 0) {
    return -2;
  }
  return -1;
}

// src/Tokenizer.ts
function outputLink(cap, link2, raw, lexer2, rules) {
  const href = link2.href;
  const title = link2.title || null;
  const text = cap[1].replace(rules.other.outputLinkReplace, "$1");
  lexer2.state.inLink = true;
  const token = {
    type: cap[0].charAt(0) === "!" ? "image" : "link",
    raw,
    href,
    title,
    text,
    tokens: lexer2.inlineTokens(text)
  };
  lexer2.state.inLink = false;
  return token;
}
function indentCodeCompensation(raw, text, rules) {
  const matchIndentToCode = raw.match(rules.other.indentCodeCompensation);
  if (matchIndentToCode === null) {
    return text;
  }
  const indentToCode = matchIndentToCode[1];
  return text.split("\n").map((node) => {
    const matchIndentInNode = node.match(rules.other.beginningSpace);
    if (matchIndentInNode === null) {
      return node;
    }
    const [indentInNode] = matchIndentInNode;
    if (indentInNode.length >= indentToCode.length) {
      return node.slice(indentToCode.length);
    }
    return node;
  }).join("\n");
}
var _Tokenizer = class {
  options;
  rules;
  // set by the lexer
  lexer;
  // set by the lexer
  constructor(options2) {
    this.options = options2 || _defaults;
  }
  space(src) {
    const cap = this.rules.block.newline.exec(src);
    if (cap && cap[0].length > 0) {
      return {
        type: "space",
        raw: cap[0]
      };
    }
  }
  code(src) {
    const cap = this.rules.block.code.exec(src);
    if (cap) {
      const text = cap[0].replace(this.rules.other.codeRemoveIndent, "");
      return {
        type: "code",
        raw: cap[0],
        codeBlockStyle: "indented",
        text: !this.options.pedantic ? rtrim(text, "\n") : text
      };
    }
  }
  fences(src) {
    const cap = this.rules.block.fences.exec(src);
    if (cap) {
      const raw = cap[0];
      const text = indentCodeCompensation(raw, cap[3] || "", this.rules);
      return {
        type: "code",
        raw,
        lang: cap[2] ? cap[2].trim().replace(this.rules.inline.anyPunctuation, "$1") : cap[2],
        text
      };
    }
  }
  heading(src) {
    const cap = this.rules.block.heading.exec(src);
    if (cap) {
      let text = cap[2].trim();
      if (this.rules.other.endingHash.test(text)) {
        const trimmed = rtrim(text, "#");
        if (this.options.pedantic) {
          text = trimmed.trim();
        } else if (!trimmed || this.rules.other.endingSpaceChar.test(trimmed)) {
          text = trimmed.trim();
        }
      }
      return {
        type: "heading",
        raw: cap[0],
        depth: cap[1].length,
        text,
        tokens: this.lexer.inline(text)
      };
    }
  }
  hr(src) {
    const cap = this.rules.block.hr.exec(src);
    if (cap) {
      return {
        type: "hr",
        raw: rtrim(cap[0], "\n")
      };
    }
  }
  blockquote(src) {
    const cap = this.rules.block.blockquote.exec(src);
    if (cap) {
      let lines = rtrim(cap[0], "\n").split("\n");
      let raw = "";
      let text = "";
      const tokens = [];
      while (lines.length > 0) {
        let inBlockquote = false;
        const currentLines = [];
        let i;
        for (i = 0; i < lines.length; i++) {
          if (this.rules.other.blockquoteStart.test(lines[i])) {
            currentLines.push(lines[i]);
            inBlockquote = true;
          } else if (!inBlockquote) {
            currentLines.push(lines[i]);
          } else {
            break;
          }
        }
        lines = lines.slice(i);
        const currentRaw = currentLines.join("\n");
        const currentText = currentRaw.replace(this.rules.other.blockquoteSetextReplace, "\n    $1").replace(this.rules.other.blockquoteSetextReplace2, "");
        raw = raw ? `${raw}
${currentRaw}` : currentRaw;
        text = text ? `${text}
${currentText}` : currentText;
        const top = this.lexer.state.top;
        this.lexer.state.top = true;
        this.lexer.blockTokens(currentText, tokens, true);
        this.lexer.state.top = top;
        if (lines.length === 0) {
          break;
        }
        const lastToken = tokens.at(-1);
        if (lastToken?.type === "code") {
          break;
        } else if (lastToken?.type === "blockquote") {
          const oldToken = lastToken;
          const newText = oldToken.raw + "\n" + lines.join("\n");
          const newToken = this.blockquote(newText);
          tokens[tokens.length - 1] = newToken;
          raw = raw.substring(0, raw.length - oldToken.raw.length) + newToken.raw;
          text = text.substring(0, text.length - oldToken.text.length) + newToken.text;
          break;
        } else if (lastToken?.type === "list") {
          const oldToken = lastToken;
          const newText = oldToken.raw + "\n" + lines.join("\n");
          const newToken = this.list(newText);
          tokens[tokens.length - 1] = newToken;
          raw = raw.substring(0, raw.length - lastToken.raw.length) + newToken.raw;
          text = text.substring(0, text.length - oldToken.raw.length) + newToken.raw;
          lines = newText.substring(tokens.at(-1).raw.length).split("\n");
          continue;
        }
      }
      return {
        type: "blockquote",
        raw,
        tokens,
        text
      };
    }
  }
  list(src) {
    let cap = this.rules.block.list.exec(src);
    if (cap) {
      let bull = cap[1].trim();
      const isordered = bull.length > 1;
      const list2 = {
        type: "list",
        raw: "",
        ordered: isordered,
        start: isordered ? +bull.slice(0, -1) : "",
        loose: false,
        items: []
      };
      bull = isordered ? `\\d{1,9}\\${bull.slice(-1)}` : `\\${bull}`;
      if (this.options.pedantic) {
        bull = isordered ? bull : "[*+-]";
      }
      const itemRegex = this.rules.other.listItemRegex(bull);
      let endsWithBlankLine = false;
      while (src) {
        let endEarly = false;
        let raw = "";
        let itemContents = "";
        if (!(cap = itemRegex.exec(src))) {
          break;
        }
        if (this.rules.block.hr.test(src)) {
          break;
        }
        raw = cap[0];
        src = src.substring(raw.length);
        let line = cap[2].split("\n", 1)[0].replace(this.rules.other.listReplaceTabs, (t) => " ".repeat(3 * t.length));
        let nextLine = src.split("\n", 1)[0];
        let blankLine = !line.trim();
        let indent = 0;
        if (this.options.pedantic) {
          indent = 2;
          itemContents = line.trimStart();
        } else if (blankLine) {
          indent = cap[1].length + 1;
        } else {
          indent = cap[2].search(this.rules.other.nonSpaceChar);
          indent = indent > 4 ? 1 : indent;
          itemContents = line.slice(indent);
          indent += cap[1].length;
        }
        if (blankLine && this.rules.other.blankLine.test(nextLine)) {
          raw += nextLine + "\n";
          src = src.substring(nextLine.length + 1);
          endEarly = true;
        }
        if (!endEarly) {
          const nextBulletRegex = this.rules.other.nextBulletRegex(indent);
          const hrRegex = this.rules.other.hrRegex(indent);
          const fencesBeginRegex = this.rules.other.fencesBeginRegex(indent);
          const headingBeginRegex = this.rules.other.headingBeginRegex(indent);
          const htmlBeginRegex = this.rules.other.htmlBeginRegex(indent);
          while (src) {
            const rawLine = src.split("\n", 1)[0];
            let nextLineWithoutTabs;
            nextLine = rawLine;
            if (this.options.pedantic) {
              nextLine = nextLine.replace(this.rules.other.listReplaceNesting, "  ");
              nextLineWithoutTabs = nextLine;
            } else {
              nextLineWithoutTabs = nextLine.replace(this.rules.other.tabCharGlobal, "    ");
            }
            if (fencesBeginRegex.test(nextLine)) {
              break;
            }
            if (headingBeginRegex.test(nextLine)) {
              break;
            }
            if (htmlBeginRegex.test(nextLine)) {
              break;
            }
            if (nextBulletRegex.test(nextLine)) {
              break;
            }
            if (hrRegex.test(nextLine)) {
              break;
            }
            if (nextLineWithoutTabs.search(this.rules.other.nonSpaceChar) >= indent || !nextLine.trim()) {
              itemContents += "\n" + nextLineWithoutTabs.slice(indent);
            } else {
              if (blankLine) {
                break;
              }
              if (line.replace(this.rules.other.tabCharGlobal, "    ").search(this.rules.other.nonSpaceChar) >= 4) {
                break;
              }
              if (fencesBeginRegex.test(line)) {
                break;
              }
              if (headingBeginRegex.test(line)) {
                break;
              }
              if (hrRegex.test(line)) {
                break;
              }
              itemContents += "\n" + nextLine;
            }
            if (!blankLine && !nextLine.trim()) {
              blankLine = true;
            }
            raw += rawLine + "\n";
            src = src.substring(rawLine.length + 1);
            line = nextLineWithoutTabs.slice(indent);
          }
        }
        if (!list2.loose) {
          if (endsWithBlankLine) {
            list2.loose = true;
          } else if (this.rules.other.doubleBlankLine.test(raw)) {
            endsWithBlankLine = true;
          }
        }
        let istask = null;
        let ischecked;
        if (this.options.gfm) {
          istask = this.rules.other.listIsTask.exec(itemContents);
          if (istask) {
            ischecked = istask[0] !== "[ ] ";
            itemContents = itemContents.replace(this.rules.other.listReplaceTask, "");
          }
        }
        list2.items.push({
          type: "list_item",
          raw,
          task: !!istask,
          checked: ischecked,
          loose: false,
          text: itemContents,
          tokens: []
        });
        list2.raw += raw;
      }
      const lastItem = list2.items.at(-1);
      if (lastItem) {
        lastItem.raw = lastItem.raw.trimEnd();
        lastItem.text = lastItem.text.trimEnd();
      } else {
        return;
      }
      list2.raw = list2.raw.trimEnd();
      for (let i = 0; i < list2.items.length; i++) {
        this.lexer.state.top = false;
        list2.items[i].tokens = this.lexer.blockTokens(list2.items[i].text, []);
        if (!list2.loose) {
          const spacers = list2.items[i].tokens.filter((t) => t.type === "space");
          const hasMultipleLineBreaks = spacers.length > 0 && spacers.some((t) => this.rules.other.anyLine.test(t.raw));
          list2.loose = hasMultipleLineBreaks;
        }
      }
      if (list2.loose) {
        for (let i = 0; i < list2.items.length; i++) {
          list2.items[i].loose = true;
        }
      }
      return list2;
    }
  }
  html(src) {
    const cap = this.rules.block.html.exec(src);
    if (cap) {
      const token = {
        type: "html",
        block: true,
        raw: cap[0],
        pre: cap[1] === "pre" || cap[1] === "script" || cap[1] === "style",
        text: cap[0]
      };
      return token;
    }
  }
  def(src) {
    const cap = this.rules.block.def.exec(src);
    if (cap) {
      const tag2 = cap[1].toLowerCase().replace(this.rules.other.multipleSpaceGlobal, " ");
      const href = cap[2] ? cap[2].replace(this.rules.other.hrefBrackets, "$1").replace(this.rules.inline.anyPunctuation, "$1") : "";
      const title = cap[3] ? cap[3].substring(1, cap[3].length - 1).replace(this.rules.inline.anyPunctuation, "$1") : cap[3];
      return {
        type: "def",
        tag: tag2,
        raw: cap[0],
        href,
        title
      };
    }
  }
  table(src) {
    const cap = this.rules.block.table.exec(src);
    if (!cap) {
      return;
    }
    if (!this.rules.other.tableDelimiter.test(cap[2])) {
      return;
    }
    const headers = splitCells(cap[1]);
    const aligns = cap[2].replace(this.rules.other.tableAlignChars, "").split("|");
    const rows = cap[3]?.trim() ? cap[3].replace(this.rules.other.tableRowBlankLine, "").split("\n") : [];
    const item = {
      type: "table",
      raw: cap[0],
      header: [],
      align: [],
      rows: []
    };
    if (headers.length !== aligns.length) {
      return;
    }
    for (const align of aligns) {
      if (this.rules.other.tableAlignRight.test(align)) {
        item.align.push("right");
      } else if (this.rules.other.tableAlignCenter.test(align)) {
        item.align.push("center");
      } else if (this.rules.other.tableAlignLeft.test(align)) {
        item.align.push("left");
      } else {
        item.align.push(null);
      }
    }
    for (let i = 0; i < headers.length; i++) {
      item.header.push({
        text: headers[i],
        tokens: this.lexer.inline(headers[i]),
        header: true,
        align: item.align[i]
      });
    }
    for (const row of rows) {
      item.rows.push(splitCells(row, item.header.length).map((cell, i) => {
        return {
          text: cell,
          tokens: this.lexer.inline(cell),
          header: false,
          align: item.align[i]
        };
      }));
    }
    return item;
  }
  lheading(src) {
    const cap = this.rules.block.lheading.exec(src);
    if (cap) {
      return {
        type: "heading",
        raw: cap[0],
        depth: cap[2].charAt(0) === "=" ? 1 : 2,
        text: cap[1],
        tokens: this.lexer.inline(cap[1])
      };
    }
  }
  paragraph(src) {
    const cap = this.rules.block.paragraph.exec(src);
    if (cap) {
      const text = cap[1].charAt(cap[1].length - 1) === "\n" ? cap[1].slice(0, -1) : cap[1];
      return {
        type: "paragraph",
        raw: cap[0],
        text,
        tokens: this.lexer.inline(text)
      };
    }
  }
  text(src) {
    const cap = this.rules.block.text.exec(src);
    if (cap) {
      return {
        type: "text",
        raw: cap[0],
        text: cap[0],
        tokens: this.lexer.inline(cap[0])
      };
    }
  }
  escape(src) {
    const cap = this.rules.inline.escape.exec(src);
    if (cap) {
      return {
        type: "escape",
        raw: cap[0],
        text: cap[1]
      };
    }
  }
  tag(src) {
    const cap = this.rules.inline.tag.exec(src);
    if (cap) {
      if (!this.lexer.state.inLink && this.rules.other.startATag.test(cap[0])) {
        this.lexer.state.inLink = true;
      } else if (this.lexer.state.inLink && this.rules.other.endATag.test(cap[0])) {
        this.lexer.state.inLink = false;
      }
      if (!this.lexer.state.inRawBlock && this.rules.other.startPreScriptTag.test(cap[0])) {
        this.lexer.state.inRawBlock = true;
      } else if (this.lexer.state.inRawBlock && this.rules.other.endPreScriptTag.test(cap[0])) {
        this.lexer.state.inRawBlock = false;
      }
      return {
        type: "html",
        raw: cap[0],
        inLink: this.lexer.state.inLink,
        inRawBlock: this.lexer.state.inRawBlock,
        block: false,
        text: cap[0]
      };
    }
  }
  link(src) {
    const cap = this.rules.inline.link.exec(src);
    if (cap) {
      const trimmedUrl = cap[2].trim();
      if (!this.options.pedantic && this.rules.other.startAngleBracket.test(trimmedUrl)) {
        if (!this.rules.other.endAngleBracket.test(trimmedUrl)) {
          return;
        }
        const rtrimSlash = rtrim(trimmedUrl.slice(0, -1), "\\");
        if ((trimmedUrl.length - rtrimSlash.length) % 2 === 0) {
          return;
        }
      } else {
        const lastParenIndex = findClosingBracket(cap[2], "()");
        if (lastParenIndex === -2) {
          return;
        }
        if (lastParenIndex > -1) {
          const start = cap[0].indexOf("!") === 0 ? 5 : 4;
          const linkLen = start + cap[1].length + lastParenIndex;
          cap[2] = cap[2].substring(0, lastParenIndex);
          cap[0] = cap[0].substring(0, linkLen).trim();
          cap[3] = "";
        }
      }
      let href = cap[2];
      let title = "";
      if (this.options.pedantic) {
        const link2 = this.rules.other.pedanticHrefTitle.exec(href);
        if (link2) {
          href = link2[1];
          title = link2[3];
        }
      } else {
        title = cap[3] ? cap[3].slice(1, -1) : "";
      }
      href = href.trim();
      if (this.rules.other.startAngleBracket.test(href)) {
        if (this.options.pedantic && !this.rules.other.endAngleBracket.test(trimmedUrl)) {
          href = href.slice(1);
        } else {
          href = href.slice(1, -1);
        }
      }
      return outputLink(cap, {
        href: href ? href.replace(this.rules.inline.anyPunctuation, "$1") : href,
        title: title ? title.replace(this.rules.inline.anyPunctuation, "$1") : title
      }, cap[0], this.lexer, this.rules);
    }
  }
  reflink(src, links) {
    let cap;
    if ((cap = this.rules.inline.reflink.exec(src)) || (cap = this.rules.inline.nolink.exec(src))) {
      const linkString = (cap[2] || cap[1]).replace(this.rules.other.multipleSpaceGlobal, " ");
      const link2 = links[linkString.toLowerCase()];
      if (!link2) {
        const text = cap[0].charAt(0);
        return {
          type: "text",
          raw: text,
          text
        };
      }
      return outputLink(cap, link2, cap[0], this.lexer, this.rules);
    }
  }
  emStrong(src, maskedSrc, prevChar = "") {
    let match = this.rules.inline.emStrongLDelim.exec(src);
    if (!match) return;
    if (match[3] && prevChar.match(this.rules.other.unicodeAlphaNumeric)) return;
    const nextChar = match[1] || match[2] || "";
    if (!nextChar || !prevChar || this.rules.inline.punctuation.exec(prevChar)) {
      const lLength = [...match[0]].length - 1;
      let rDelim, rLength, delimTotal = lLength, midDelimTotal = 0;
      const endReg = match[0][0] === "*" ? this.rules.inline.emStrongRDelimAst : this.rules.inline.emStrongRDelimUnd;
      endReg.lastIndex = 0;
      maskedSrc = maskedSrc.slice(-1 * src.length + lLength);
      while ((match = endReg.exec(maskedSrc)) != null) {
        rDelim = match[1] || match[2] || match[3] || match[4] || match[5] || match[6];
        if (!rDelim) continue;
        rLength = [...rDelim].length;
        if (match[3] || match[4]) {
          delimTotal += rLength;
          continue;
        } else if (match[5] || match[6]) {
          if (lLength % 3 && !((lLength + rLength) % 3)) {
            midDelimTotal += rLength;
            continue;
          }
        }
        delimTotal -= rLength;
        if (delimTotal > 0) continue;
        rLength = Math.min(rLength, rLength + delimTotal + midDelimTotal);
        const lastCharLength = [...match[0]][0].length;
        const raw = src.slice(0, lLength + match.index + lastCharLength + rLength);
        if (Math.min(lLength, rLength) % 2) {
          const text2 = raw.slice(1, -1);
          return {
            type: "em",
            raw,
            text: text2,
            tokens: this.lexer.inlineTokens(text2)
          };
        }
        const text = raw.slice(2, -2);
        return {
          type: "strong",
          raw,
          text,
          tokens: this.lexer.inlineTokens(text)
        };
      }
    }
  }
  codespan(src) {
    const cap = this.rules.inline.code.exec(src);
    if (cap) {
      let text = cap[2].replace(this.rules.other.newLineCharGlobal, " ");
      const hasNonSpaceChars = this.rules.other.nonSpaceChar.test(text);
      const hasSpaceCharsOnBothEnds = this.rules.other.startingSpaceChar.test(text) && this.rules.other.endingSpaceChar.test(text);
      if (hasNonSpaceChars && hasSpaceCharsOnBothEnds) {
        text = text.substring(1, text.length - 1);
      }
      return {
        type: "codespan",
        raw: cap[0],
        text
      };
    }
  }
  br(src) {
    const cap = this.rules.inline.br.exec(src);
    if (cap) {
      return {
        type: "br",
        raw: cap[0]
      };
    }
  }
  del(src) {
    const cap = this.rules.inline.del.exec(src);
    if (cap) {
      return {
        type: "del",
        raw: cap[0],
        text: cap[2],
        tokens: this.lexer.inlineTokens(cap[2])
      };
    }
  }
  autolink(src) {
    const cap = this.rules.inline.autolink.exec(src);
    if (cap) {
      let text, href;
      if (cap[2] === "@") {
        text = cap[1];
        href = "mailto:" + text;
      } else {
        text = cap[1];
        href = text;
      }
      return {
        type: "link",
        raw: cap[0],
        text,
        href,
        tokens: [
          {
            type: "text",
            raw: text,
            text
          }
        ]
      };
    }
  }
  url(src) {
    let cap;
    if (cap = this.rules.inline.url.exec(src)) {
      let text, href;
      if (cap[2] === "@") {
        text = cap[0];
        href = "mailto:" + text;
      } else {
        let prevCapZero;
        do {
          prevCapZero = cap[0];
          cap[0] = this.rules.inline._backpedal.exec(cap[0])?.[0] ?? "";
        } while (prevCapZero !== cap[0]);
        text = cap[0];
        if (cap[1] === "www.") {
          href = "http://" + cap[0];
        } else {
          href = cap[0];
        }
      }
      return {
        type: "link",
        raw: cap[0],
        text,
        href,
        tokens: [
          {
            type: "text",
            raw: text,
            text
          }
        ]
      };
    }
  }
  inlineText(src) {
    const cap = this.rules.inline.text.exec(src);
    if (cap) {
      const escaped = this.lexer.state.inRawBlock;
      return {
        type: "text",
        raw: cap[0],
        text: cap[0],
        escaped
      };
    }
  }
};

// src/Lexer.ts
var _Lexer = class __Lexer {
  tokens;
  options;
  state;
  tokenizer;
  inlineQueue;
  constructor(options2) {
    this.tokens = [];
    this.tokens.links = /* @__PURE__ */ Object.create(null);
    this.options = options2 || _defaults;
    this.options.tokenizer = this.options.tokenizer || new _Tokenizer();
    this.tokenizer = this.options.tokenizer;
    this.tokenizer.options = this.options;
    this.tokenizer.lexer = this;
    this.inlineQueue = [];
    this.state = {
      inLink: false,
      inRawBlock: false,
      top: true
    };
    const rules = {
      other,
      block: block.normal,
      inline: inline.normal
    };
    if (this.options.pedantic) {
      rules.block = block.pedantic;
      rules.inline = inline.pedantic;
    } else if (this.options.gfm) {
      rules.block = block.gfm;
      if (this.options.breaks) {
        rules.inline = inline.breaks;
      } else {
        rules.inline = inline.gfm;
      }
    }
    this.tokenizer.rules = rules;
  }
  /**
   * Expose Rules
   */
  static get rules() {
    return {
      block,
      inline
    };
  }
  /**
   * Static Lex Method
   */
  static lex(src, options2) {
    const lexer2 = new __Lexer(options2);
    return lexer2.lex(src);
  }
  /**
   * Static Lex Inline Method
   */
  static lexInline(src, options2) {
    const lexer2 = new __Lexer(options2);
    return lexer2.inlineTokens(src);
  }
  /**
   * Preprocessing
   */
  lex(src) {
    src = src.replace(other.carriageReturn, "\n");
    this.blockTokens(src, this.tokens);
    for (let i = 0; i < this.inlineQueue.length; i++) {
      const next = this.inlineQueue[i];
      this.inlineTokens(next.src, next.tokens);
    }
    this.inlineQueue = [];
    return this.tokens;
  }
  blockTokens(src, tokens = [], lastParagraphClipped = false) {
    if (this.options.pedantic) {
      src = src.replace(other.tabCharGlobal, "    ").replace(other.spaceLine, "");
    }
    while (src) {
      let token;
      if (this.options.extensions?.block?.some((extTokenizer) => {
        if (token = extTokenizer.call({ lexer: this }, src, tokens)) {
          src = src.substring(token.raw.length);
          tokens.push(token);
          return true;
        }
        return false;
      })) {
        continue;
      }
      if (token = this.tokenizer.space(src)) {
        src = src.substring(token.raw.length);
        const lastToken = tokens.at(-1);
        if (token.raw.length === 1 && lastToken !== void 0) {
          lastToken.raw += "\n";
        } else {
          tokens.push(token);
        }
        continue;
      }
      if (token = this.tokenizer.code(src)) {
        src = src.substring(token.raw.length);
        const lastToken = tokens.at(-1);
        if (lastToken?.type === "paragraph" || lastToken?.type === "text") {
          lastToken.raw += "\n" + token.raw;
          lastToken.text += "\n" + token.text;
          this.inlineQueue.at(-1).src = lastToken.text;
        } else {
          tokens.push(token);
        }
        continue;
      }
      if (token = this.tokenizer.fences(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.heading(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.hr(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.blockquote(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.list(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.html(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.def(src)) {
        src = src.substring(token.raw.length);
        const lastToken = tokens.at(-1);
        if (lastToken?.type === "paragraph" || lastToken?.type === "text") {
          lastToken.raw += "\n" + token.raw;
          lastToken.text += "\n" + token.raw;
          this.inlineQueue.at(-1).src = lastToken.text;
        } else if (!this.tokens.links[token.tag]) {
          this.tokens.links[token.tag] = {
            href: token.href,
            title: token.title
          };
        }
        continue;
      }
      if (token = this.tokenizer.table(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.lheading(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      let cutSrc = src;
      if (this.options.extensions?.startBlock) {
        let startIndex = Infinity;
        const tempSrc = src.slice(1);
        let tempStart;
        this.options.extensions.startBlock.forEach((getStartIndex) => {
          tempStart = getStartIndex.call({ lexer: this }, tempSrc);
          if (typeof tempStart === "number" && tempStart >= 0) {
            startIndex = Math.min(startIndex, tempStart);
          }
        });
        if (startIndex < Infinity && startIndex >= 0) {
          cutSrc = src.substring(0, startIndex + 1);
        }
      }
      if (this.state.top && (token = this.tokenizer.paragraph(cutSrc))) {
        const lastToken = tokens.at(-1);
        if (lastParagraphClipped && lastToken?.type === "paragraph") {
          lastToken.raw += "\n" + token.raw;
          lastToken.text += "\n" + token.text;
          this.inlineQueue.pop();
          this.inlineQueue.at(-1).src = lastToken.text;
        } else {
          tokens.push(token);
        }
        lastParagraphClipped = cutSrc.length !== src.length;
        src = src.substring(token.raw.length);
        continue;
      }
      if (token = this.tokenizer.text(src)) {
        src = src.substring(token.raw.length);
        const lastToken = tokens.at(-1);
        if (lastToken?.type === "text") {
          lastToken.raw += "\n" + token.raw;
          lastToken.text += "\n" + token.text;
          this.inlineQueue.pop();
          this.inlineQueue.at(-1).src = lastToken.text;
        } else {
          tokens.push(token);
        }
        continue;
      }
      if (src) {
        const errMsg = "Infinite loop on byte: " + src.charCodeAt(0);
        if (this.options.silent) {
          console.error(errMsg);
          break;
        } else {
          throw new Error(errMsg);
        }
      }
    }
    this.state.top = true;
    return tokens;
  }
  inline(src, tokens = []) {
    this.inlineQueue.push({ src, tokens });
    return tokens;
  }
  /**
   * Lexing/Compiling
   */
  inlineTokens(src, tokens = []) {
    let maskedSrc = src;
    let match = null;
    if (this.tokens.links) {
      const links = Object.keys(this.tokens.links);
      if (links.length > 0) {
        while ((match = this.tokenizer.rules.inline.reflinkSearch.exec(maskedSrc)) != null) {
          if (links.includes(match[0].slice(match[0].lastIndexOf("[") + 1, -1))) {
            maskedSrc = maskedSrc.slice(0, match.index) + "[" + "a".repeat(match[0].length - 2) + "]" + maskedSrc.slice(this.tokenizer.rules.inline.reflinkSearch.lastIndex);
          }
        }
      }
    }
    while ((match = this.tokenizer.rules.inline.anyPunctuation.exec(maskedSrc)) != null) {
      maskedSrc = maskedSrc.slice(0, match.index) + "++" + maskedSrc.slice(this.tokenizer.rules.inline.anyPunctuation.lastIndex);
    }
    while ((match = this.tokenizer.rules.inline.blockSkip.exec(maskedSrc)) != null) {
      maskedSrc = maskedSrc.slice(0, match.index) + "[" + "a".repeat(match[0].length - 2) + "]" + maskedSrc.slice(this.tokenizer.rules.inline.blockSkip.lastIndex);
    }
    let keepPrevChar = false;
    let prevChar = "";
    while (src) {
      if (!keepPrevChar) {
        prevChar = "";
      }
      keepPrevChar = false;
      let token;
      if (this.options.extensions?.inline?.some((extTokenizer) => {
        if (token = extTokenizer.call({ lexer: this }, src, tokens)) {
          src = src.substring(token.raw.length);
          tokens.push(token);
          return true;
        }
        return false;
      })) {
        continue;
      }
      if (token = this.tokenizer.escape(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.tag(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.link(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.reflink(src, this.tokens.links)) {
        src = src.substring(token.raw.length);
        const lastToken = tokens.at(-1);
        if (token.type === "text" && lastToken?.type === "text") {
          lastToken.raw += token.raw;
          lastToken.text += token.text;
        } else {
          tokens.push(token);
        }
        continue;
      }
      if (token = this.tokenizer.emStrong(src, maskedSrc, prevChar)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.codespan(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.br(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.del(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (token = this.tokenizer.autolink(src)) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      if (!this.state.inLink && (token = this.tokenizer.url(src))) {
        src = src.substring(token.raw.length);
        tokens.push(token);
        continue;
      }
      let cutSrc = src;
      if (this.options.extensions?.startInline) {
        let startIndex = Infinity;
        const tempSrc = src.slice(1);
        let tempStart;
        this.options.extensions.startInline.forEach((getStartIndex) => {
          tempStart = getStartIndex.call({ lexer: this }, tempSrc);
          if (typeof tempStart === "number" && tempStart >= 0) {
            startIndex = Math.min(startIndex, tempStart);
          }
        });
        if (startIndex < Infinity && startIndex >= 0) {
          cutSrc = src.substring(0, startIndex + 1);
        }
      }
      if (token = this.tokenizer.inlineText(cutSrc)) {
        src = src.substring(token.raw.length);
        if (token.raw.slice(-1) !== "_") {
          prevChar = token.raw.slice(-1);
        }
        keepPrevChar = true;
        const lastToken = tokens.at(-1);
        if (lastToken?.type === "text") {
          lastToken.raw += token.raw;
          lastToken.text += token.text;
        } else {
          tokens.push(token);
        }
        continue;
      }
      if (src) {
        const errMsg = "Infinite loop on byte: " + src.charCodeAt(0);
        if (this.options.silent) {
          console.error(errMsg);
          break;
        } else {
          throw new Error(errMsg);
        }
      }
    }
    return tokens;
  }
};

// src/Renderer.ts
var _Renderer = class {
  options;
  parser;
  // set by the parser
  constructor(options2) {
    this.options = options2 || _defaults;
  }
  space(token) {
    return "";
  }
  code({ text, lang, escaped }) {
    const langString = (lang || "").match(other.notSpaceStart)?.[0];
    const code = text.replace(other.endingNewline, "") + "\n";
    if (!langString) {
      return "<pre><code>" + (escaped ? code : escape2(code, true)) + "</code></pre>\n";
    }
    return '<pre><code class="language-' + escape2(langString) + '">' + (escaped ? code : escape2(code, true)) + "</code></pre>\n";
  }
  blockquote({ tokens }) {
    const body = this.parser.parse(tokens);
    return `<blockquote>
${body}</blockquote>
`;
  }
  html({ text }) {
    return text;
  }
  heading({ tokens, depth }) {
    return `<h${depth}>${this.parser.parseInline(tokens)}</h${depth}>
`;
  }
  hr(token) {
    return "<hr>\n";
  }
  list(token) {
    const ordered = token.ordered;
    const start = token.start;
    let body = "";
    for (let j = 0; j < token.items.length; j++) {
      const item = token.items[j];
      body += this.listitem(item);
    }
    const type = ordered ? "ol" : "ul";
    const startAttr = ordered && start !== 1 ? ' start="' + start + '"' : "";
    return "<" + type + startAttr + ">\n" + body + "</" + type + ">\n";
  }
  listitem(item) {
    let itemBody = "";
    if (item.task) {
      const checkbox = this.checkbox({ checked: !!item.checked });
      if (item.loose) {
        if (item.tokens[0]?.type === "paragraph") {
          item.tokens[0].text = checkbox + " " + item.tokens[0].text;
          if (item.tokens[0].tokens && item.tokens[0].tokens.length > 0 && item.tokens[0].tokens[0].type === "text") {
            item.tokens[0].tokens[0].text = checkbox + " " + escape2(item.tokens[0].tokens[0].text);
            item.tokens[0].tokens[0].escaped = true;
          }
        } else {
          item.tokens.unshift({
            type: "text",
            raw: checkbox + " ",
            text: checkbox + " ",
            escaped: true
          });
        }
      } else {
        itemBody += checkbox + " ";
      }
    }
    itemBody += this.parser.parse(item.tokens, !!item.loose);
    return `<li>${itemBody}</li>
`;
  }
  checkbox({ checked }) {
    return "<input " + (checked ? 'checked="" ' : "") + 'disabled="" type="checkbox">';
  }
  paragraph({ tokens }) {
    return `<p>${this.parser.parseInline(tokens)}</p>
`;
  }
  table(token) {
    let header = "";
    let cell = "";
    for (let j = 0; j < token.header.length; j++) {
      cell += this.tablecell(token.header[j]);
    }
    header += this.tablerow({ text: cell });
    let body = "";
    for (let j = 0; j < token.rows.length; j++) {
      const row = token.rows[j];
      cell = "";
      for (let k = 0; k < row.length; k++) {
        cell += this.tablecell(row[k]);
      }
      body += this.tablerow({ text: cell });
    }
    if (body) body = `<tbody>${body}</tbody>`;
    return "<table>\n<thead>\n" + header + "</thead>\n" + body + "</table>\n";
  }
  tablerow({ text }) {
    return `<tr>
${text}</tr>
`;
  }
  tablecell(token) {
    const content = this.parser.parseInline(token.tokens);
    const type = token.header ? "th" : "td";
    const tag2 = token.align ? `<${type} align="${token.align}">` : `<${type}>`;
    return tag2 + content + `</${type}>
`;
  }
  /**
   * span level renderer
   */
  strong({ tokens }) {
    return `<strong>${this.parser.parseInline(tokens)}</strong>`;
  }
  em({ tokens }) {
    return `<em>${this.parser.parseInline(tokens)}</em>`;
  }
  codespan({ text }) {
    return `<code>${escape2(text, true)}</code>`;
  }
  br(token) {
    return "<br>";
  }
  del({ tokens }) {
    return `<del>${this.parser.parseInline(tokens)}</del>`;
  }
  link({ href, title, tokens }) {
    const text = this.parser.parseInline(tokens);
    const cleanHref = cleanUrl(href);
    if (cleanHref === null) {
      return text;
    }
    href = cleanHref;
    let out = '<a href="' + href + '"';
    if (title) {
      out += ' title="' + escape2(title) + '"';
    }
    out += ">" + text + "</a>";
    return out;
  }
  image({ href, title, text, tokens }) {
    if (tokens) {
      text = this.parser.parseInline(tokens, this.parser.textRenderer);
    }
    const cleanHref = cleanUrl(href);
    if (cleanHref === null) {
      return escape2(text);
    }
    href = cleanHref;
    let out = `<img src="${href}" alt="${text}"`;
    if (title) {
      out += ` title="${escape2(title)}"`;
    }
    out += ">";
    return out;
  }
  text(token) {
    return "tokens" in token && token.tokens ? this.parser.parseInline(token.tokens) : "escaped" in token && token.escaped ? token.text : escape2(token.text);
  }
};

// src/TextRenderer.ts
var _TextRenderer = class {
  // no need for block level renderers
  strong({ text }) {
    return text;
  }
  em({ text }) {
    return text;
  }
  codespan({ text }) {
    return text;
  }
  del({ text }) {
    return text;
  }
  html({ text }) {
    return text;
  }
  text({ text }) {
    return text;
  }
  link({ text }) {
    return "" + text;
  }
  image({ text }) {
    return "" + text;
  }
  br() {
    return "";
  }
};

// src/Parser.ts
var _Parser = class __Parser {
  options;
  renderer;
  textRenderer;
  constructor(options2) {
    this.options = options2 || _defaults;
    this.options.renderer = this.options.renderer || new _Renderer();
    this.renderer = this.options.renderer;
    this.renderer.options = this.options;
    this.renderer.parser = this;
    this.textRenderer = new _TextRenderer();
  }
  /**
   * Static Parse Method
   */
  static parse(tokens, options2) {
    const parser2 = new __Parser(options2);
    return parser2.parse(tokens);
  }
  /**
   * Static Parse Inline Method
   */
  static parseInline(tokens, options2) {
    const parser2 = new __Parser(options2);
    return parser2.parseInline(tokens);
  }
  /**
   * Parse Loop
   */
  parse(tokens, top = true) {
    let out = "";
    for (let i = 0; i < tokens.length; i++) {
      const anyToken = tokens[i];
      if (this.options.extensions?.renderers?.[anyToken.type]) {
        const genericToken = anyToken;
        const ret = this.options.extensions.renderers[genericToken.type].call({ parser: this }, genericToken);
        if (ret !== false || !["space", "hr", "heading", "code", "table", "blockquote", "list", "html", "paragraph", "text"].includes(genericToken.type)) {
          out += ret || "";
          continue;
        }
      }
      const token = anyToken;
      switch (token.type) {
        case "space": {
          out += this.renderer.space(token);
          continue;
        }
        case "hr": {
          out += this.renderer.hr(token);
          continue;
        }
        case "heading": {
          out += this.renderer.heading(token);
          continue;
        }
        case "code": {
          out += this.renderer.code(token);
          continue;
        }
        case "table": {
          out += this.renderer.table(token);
          continue;
        }
        case "blockquote": {
          out += this.renderer.blockquote(token);
          continue;
        }
        case "list": {
          out += this.renderer.list(token);
          continue;
        }
        case "html": {
          out += this.renderer.html(token);
          continue;
        }
        case "paragraph": {
          out += this.renderer.paragraph(token);
          continue;
        }
        case "text": {
          let textToken = token;
          let body = this.renderer.text(textToken);
          while (i + 1 < tokens.length && tokens[i + 1].type === "text") {
            textToken = tokens[++i];
            body += "\n" + this.renderer.text(textToken);
          }
          if (top) {
            out += this.renderer.paragraph({
              type: "paragraph",
              raw: body,
              text: body,
              tokens: [{ type: "text", raw: body, text: body, escaped: true }]
            });
          } else {
            out += body;
          }
          continue;
        }
        default: {
          const errMsg = 'Token with "' + token.type + '" type was not found.';
          if (this.options.silent) {
            console.error(errMsg);
            return "";
          } else {
            throw new Error(errMsg);
          }
        }
      }
    }
    return out;
  }
  /**
   * Parse Inline Tokens
   */
  parseInline(tokens, renderer = this.renderer) {
    let out = "";
    for (let i = 0; i < tokens.length; i++) {
      const anyToken = tokens[i];
      if (this.options.extensions?.renderers?.[anyToken.type]) {
        const ret = this.options.extensions.renderers[anyToken.type].call({ parser: this }, anyToken);
        if (ret !== false || !["escape", "html", "link", "image", "strong", "em", "codespan", "br", "del", "text"].includes(anyToken.type)) {
          out += ret || "";
          continue;
        }
      }
      const token = anyToken;
      switch (token.type) {
        case "escape": {
          out += renderer.text(token);
          break;
        }
        case "html": {
          out += renderer.html(token);
          break;
        }
        case "link": {
          out += renderer.link(token);
          break;
        }
        case "image": {
          out += renderer.image(token);
          break;
        }
        case "strong": {
          out += renderer.strong(token);
          break;
        }
        case "em": {
          out += renderer.em(token);
          break;
        }
        case "codespan": {
          out += renderer.codespan(token);
          break;
        }
        case "br": {
          out += renderer.br(token);
          break;
        }
        case "del": {
          out += renderer.del(token);
          break;
        }
        case "text": {
          out += renderer.text(token);
          break;
        }
        default: {
          const errMsg = 'Token with "' + token.type + '" type was not found.';
          if (this.options.silent) {
            console.error(errMsg);
            return "";
          } else {
            throw new Error(errMsg);
          }
        }
      }
    }
    return out;
  }
};

// src/Hooks.ts
var _Hooks = class {
  options;
  block;
  constructor(options2) {
    this.options = options2 || _defaults;
  }
  static passThroughHooks = /* @__PURE__ */ new Set([
    "preprocess",
    "postprocess",
    "processAllTokens"
  ]);
  /**
   * Process markdown before marked
   */
  preprocess(markdown) {
    return markdown;
  }
  /**
   * Process HTML after marked is finished
   */
  postprocess(html2) {
    return html2;
  }
  /**
   * Process all tokens before walk tokens
   */
  processAllTokens(tokens) {
    return tokens;
  }
  /**
   * Provide function to tokenize markdown
   */
  provideLexer() {
    return this.block ? _Lexer.lex : _Lexer.lexInline;
  }
  /**
   * Provide function to parse tokens
   */
  provideParser() {
    return this.block ? _Parser.parse : _Parser.parseInline;
  }
};

// src/Instance.ts
var Marked = class {
  defaults = _getDefaults();
  options = this.setOptions;
  parse = this.parseMarkdown(true);
  parseInline = this.parseMarkdown(false);
  Parser = _Parser;
  Renderer = _Renderer;
  TextRenderer = _TextRenderer;
  Lexer = _Lexer;
  Tokenizer = _Tokenizer;
  Hooks = _Hooks;
  constructor(...args) {
    this.use(...args);
  }
  /**
   * Run callback for every token
   */
  walkTokens(tokens, callback) {
    let values = [];
    for (const token of tokens) {
      values = values.concat(callback.call(this, token));
      switch (token.type) {
        case "table": {
          const tableToken = token;
          for (const cell of tableToken.header) {
            values = values.concat(this.walkTokens(cell.tokens, callback));
          }
          for (const row of tableToken.rows) {
            for (const cell of row) {
              values = values.concat(this.walkTokens(cell.tokens, callback));
            }
          }
          break;
        }
        case "list": {
          const listToken = token;
          values = values.concat(this.walkTokens(listToken.items, callback));
          break;
        }
        default: {
          const genericToken = token;
          if (this.defaults.extensions?.childTokens?.[genericToken.type]) {
            this.defaults.extensions.childTokens[genericToken.type].forEach((childTokens) => {
              const tokens2 = genericToken[childTokens].flat(Infinity);
              values = values.concat(this.walkTokens(tokens2, callback));
            });
          } else if (genericToken.tokens) {
            values = values.concat(this.walkTokens(genericToken.tokens, callback));
          }
        }
      }
    }
    return values;
  }
  use(...args) {
    const extensions = this.defaults.extensions || { renderers: {}, childTokens: {} };
    args.forEach((pack) => {
      const opts = { ...pack };
      opts.async = this.defaults.async || opts.async || false;
      if (pack.extensions) {
        pack.extensions.forEach((ext) => {
          if (!ext.name) {
            throw new Error("extension name required");
          }
          if ("renderer" in ext) {
            const prevRenderer = extensions.renderers[ext.name];
            if (prevRenderer) {
              extensions.renderers[ext.name] = function(...args2) {
                let ret = ext.renderer.apply(this, args2);
                if (ret === false) {
                  ret = prevRenderer.apply(this, args2);
                }
                return ret;
              };
            } else {
              extensions.renderers[ext.name] = ext.renderer;
            }
          }
          if ("tokenizer" in ext) {
            if (!ext.level || ext.level !== "block" && ext.level !== "inline") {
              throw new Error("extension level must be 'block' or 'inline'");
            }
            const extLevel = extensions[ext.level];
            if (extLevel) {
              extLevel.unshift(ext.tokenizer);
            } else {
              extensions[ext.level] = [ext.tokenizer];
            }
            if (ext.start) {
              if (ext.level === "block") {
                if (extensions.startBlock) {
                  extensions.startBlock.push(ext.start);
                } else {
                  extensions.startBlock = [ext.start];
                }
              } else if (ext.level === "inline") {
                if (extensions.startInline) {
                  extensions.startInline.push(ext.start);
                } else {
                  extensions.startInline = [ext.start];
                }
              }
            }
          }
          if ("childTokens" in ext && ext.childTokens) {
            extensions.childTokens[ext.name] = ext.childTokens;
          }
        });
        opts.extensions = extensions;
      }
      if (pack.renderer) {
        const renderer = this.defaults.renderer || new _Renderer(this.defaults);
        for (const prop in pack.renderer) {
          if (!(prop in renderer)) {
            throw new Error(`renderer '${prop}' does not exist`);
          }
          if (["options", "parser"].includes(prop)) {
            continue;
          }
          const rendererProp = prop;
          const rendererFunc = pack.renderer[rendererProp];
          const prevRenderer = renderer[rendererProp];
          renderer[rendererProp] = (...args2) => {
            let ret = rendererFunc.apply(renderer, args2);
            if (ret === false) {
              ret = prevRenderer.apply(renderer, args2);
            }
            return ret || "";
          };
        }
        opts.renderer = renderer;
      }
      if (pack.tokenizer) {
        const tokenizer = this.defaults.tokenizer || new _Tokenizer(this.defaults);
        for (const prop in pack.tokenizer) {
          if (!(prop in tokenizer)) {
            throw new Error(`tokenizer '${prop}' does not exist`);
          }
          if (["options", "rules", "lexer"].includes(prop)) {
            continue;
          }
          const tokenizerProp = prop;
          const tokenizerFunc = pack.tokenizer[tokenizerProp];
          const prevTokenizer = tokenizer[tokenizerProp];
          tokenizer[tokenizerProp] = (...args2) => {
            let ret = tokenizerFunc.apply(tokenizer, args2);
            if (ret === false) {
              ret = prevTokenizer.apply(tokenizer, args2);
            }
            return ret;
          };
        }
        opts.tokenizer = tokenizer;
      }
      if (pack.hooks) {
        const hooks = this.defaults.hooks || new _Hooks();
        for (const prop in pack.hooks) {
          if (!(prop in hooks)) {
            throw new Error(`hook '${prop}' does not exist`);
          }
          if (["options", "block"].includes(prop)) {
            continue;
          }
          const hooksProp = prop;
          const hooksFunc = pack.hooks[hooksProp];
          const prevHook = hooks[hooksProp];
          if (_Hooks.passThroughHooks.has(prop)) {
            hooks[hooksProp] = (arg) => {
              if (this.defaults.async) {
                return Promise.resolve(hooksFunc.call(hooks, arg)).then((ret2) => {
                  return prevHook.call(hooks, ret2);
                });
              }
              const ret = hooksFunc.call(hooks, arg);
              return prevHook.call(hooks, ret);
            };
          } else {
            hooks[hooksProp] = (...args2) => {
              let ret = hooksFunc.apply(hooks, args2);
              if (ret === false) {
                ret = prevHook.apply(hooks, args2);
              }
              return ret;
            };
          }
        }
        opts.hooks = hooks;
      }
      if (pack.walkTokens) {
        const walkTokens2 = this.defaults.walkTokens;
        const packWalktokens = pack.walkTokens;
        opts.walkTokens = function(token) {
          let values = [];
          values.push(packWalktokens.call(this, token));
          if (walkTokens2) {
            values = values.concat(walkTokens2.call(this, token));
          }
          return values;
        };
      }
      this.defaults = { ...this.defaults, ...opts };
    });
    return this;
  }
  setOptions(opt) {
    this.defaults = { ...this.defaults, ...opt };
    return this;
  }
  lexer(src, options2) {
    return _Lexer.lex(src, options2 ?? this.defaults);
  }
  parser(tokens, options2) {
    return _Parser.parse(tokens, options2 ?? this.defaults);
  }
  parseMarkdown(blockType) {
    const parse2 = (src, options2) => {
      const origOpt = { ...options2 };
      const opt = { ...this.defaults, ...origOpt };
      const throwError = this.onError(!!opt.silent, !!opt.async);
      if (this.defaults.async === true && origOpt.async === false) {
        return throwError(new Error("marked(): The async option was set to true by an extension. Remove async: false from the parse options object to return a Promise."));
      }
      if (typeof src === "undefined" || src === null) {
        return throwError(new Error("marked(): input parameter is undefined or null"));
      }
      if (typeof src !== "string") {
        return throwError(new Error("marked(): input parameter is of type " + Object.prototype.toString.call(src) + ", string expected"));
      }
      if (opt.hooks) {
        opt.hooks.options = opt;
        opt.hooks.block = blockType;
      }
      const lexer2 = opt.hooks ? opt.hooks.provideLexer() : blockType ? _Lexer.lex : _Lexer.lexInline;
      const parser2 = opt.hooks ? opt.hooks.provideParser() : blockType ? _Parser.parse : _Parser.parseInline;
      if (opt.async) {
        return Promise.resolve(opt.hooks ? opt.hooks.preprocess(src) : src).then((src2) => lexer2(src2, opt)).then((tokens) => opt.hooks ? opt.hooks.processAllTokens(tokens) : tokens).then((tokens) => opt.walkTokens ? Promise.all(this.walkTokens(tokens, opt.walkTokens)).then(() => tokens) : tokens).then((tokens) => parser2(tokens, opt)).then((html2) => opt.hooks ? opt.hooks.postprocess(html2) : html2).catch(throwError);
      }
      try {
        if (opt.hooks) {
          src = opt.hooks.preprocess(src);
        }
        let tokens = lexer2(src, opt);
        if (opt.hooks) {
          tokens = opt.hooks.processAllTokens(tokens);
        }
        if (opt.walkTokens) {
          this.walkTokens(tokens, opt.walkTokens);
        }
        let html2 = parser2(tokens, opt);
        if (opt.hooks) {
          html2 = opt.hooks.postprocess(html2);
        }
        return html2;
      } catch (e) {
        return throwError(e);
      }
    };
    return parse2;
  }
  onError(silent, async) {
    return (e) => {
      e.message += "\nPlease report this to https://github.com/markedjs/marked.";
      if (silent) {
        const msg = "<p>An error occurred:</p><pre>" + escape2(e.message + "", true) + "</pre>";
        if (async) {
          return Promise.resolve(msg);
        }
        return msg;
      }
      if (async) {
        return Promise.reject(e);
      }
      throw e;
    };
  }
};

// src/marked.ts
var markedInstance = new Marked();
function marked(src, opt) {
  return markedInstance.parse(src, opt);
}
marked.options = marked.setOptions = function(options2) {
  markedInstance.setOptions(options2);
  marked.defaults = markedInstance.defaults;
  changeDefaults(marked.defaults);
  return marked;
};
marked.getDefaults = _getDefaults;
marked.defaults = _defaults;
marked.use = function(...args) {
  markedInstance.use(...args);
  marked.defaults = markedInstance.defaults;
  changeDefaults(marked.defaults);
  return marked;
};
marked.walkTokens = function(tokens, callback) {
  return markedInstance.walkTokens(tokens, callback);
};
marked.parseInline = markedInstance.parseInline;
marked.Parser = _Parser;
marked.parser = _Parser.parse;
marked.Renderer = _Renderer;
marked.TextRenderer = _TextRenderer;
marked.Lexer = _Lexer;
marked.lexer = _Lexer.lex;
marked.Tokenizer = _Tokenizer;
marked.Hooks = _Hooks;
marked.parse = marked;
var options = marked.options;
var setOptions = marked.setOptions;
var use = marked.use;
var walkTokens = marked.walkTokens;
var parseInline = marked.parseInline;
var parse = (/* unused pure expression or super */ null && (marked));
var parser = _Parser.parse;
var lexer = _Lexer.lex;

//# sourceMappingURL=marked.esm.js.map

;// ./src/toolErrors.ts
var ToolErrorCategory;
(function (ToolErrorCategory) {
    ToolErrorCategory["Validation"] = "validation";
    ToolErrorCategory["Timeout"] = "timeout";
    ToolErrorCategory["Cancellation"] = "cancellation";
    ToolErrorCategory["Repetition"] = "repetition";
    ToolErrorCategory["OfficeHost"] = "office_host";
    ToolErrorCategory["Network"] = "network";
    ToolErrorCategory["Provider"] = "provider";
    ToolErrorCategory["Protocol"] = "protocol";
    ToolErrorCategory["Internal"] = "internal";
})(ToolErrorCategory || (ToolErrorCategory = {}));
const trustedToolError = Symbol("trustedToolError");
const SAFE_ERROR_NAMES = new Set(["Error", "TypeError", "RangeError", "SyntaxError", "ReferenceError", "URIError", "EvalError", "DOMException", "AbortError", "TimeoutError"]);
const SAFE_CODE = /^[A-Za-z0-9]{1,32}$/;
const SAFE_TOKEN = /^[A-Za-z0-9._-]{1,64}$/;
const SAFE_CAPABILITY_TEXT = /^[\x20-\x7e]{1,512}$/;
const FIXED_MESSAGES = {
    validation: "The tool request was invalid.",
    timeout: "The operation exceeded its time limit.",
    cancellation: "The running Hermes task was stopped.",
    repetition: "The repeated tool request was blocked.",
    office_host: "PowerPoint could not complete the requested operation.",
    network: "Could not reach the model endpoint.",
    provider: "The model provider returned an error.",
    protocol: "The model response could not be processed.",
    internal: "An internal Hermes error occurred.",
};
function messageOf(value) { if (typeof value === "string")
    return value; if (value instanceof Error)
    return value.message || value.name; if (value && typeof value === "object" && typeof value.message === "string")
    return String(value.message); return "Operation failed."; }
function isCategory(value) { return typeof value === "string" && Object.values(ToolErrorCategory).includes(value); }
function safeName(value) { return typeof value === "string" && SAFE_ERROR_NAMES.has(value) ? value : undefined; }
function safeNumber(value) { return typeof value === "number" && Number.isFinite(value) ? value : undefined; }
function safeCode(value) { return typeof value === "string" && SAFE_CODE.test(value) ? value : undefined; }
function safeToken(value) { return typeof value === "string" && SAFE_TOKEN.test(value) ? value : undefined; }
function sanitizedDetails(record, trusted, allowCapabilityMetadata) {
    const source = record?.details && typeof record.details === "object" && !Array.isArray(record.details)
        ? record.details
        : undefined;
    const result = {};
    const code = safeCode(source?.code) ?? safeCode(record?.code);
    if (code)
        result.code = code;
    if (!trusted) {
        if (allowCapabilityMetadata && typeof source?.available === "boolean")
            result.available = source.available;
        if (allowCapabilityMetadata && typeof source?.requiredApi === "string" && SAFE_CAPABILITY_TEXT.test(source.requiredApi))
            result.requiredApi = source.requiredApi;
        return Object.keys(result).length ? result : undefined;
    }
    if (typeof source?.available === "boolean")
        result.available = source.available;
    if (typeof source?.requiredApi === "string" && SAFE_CAPABILITY_TEXT.test(source.requiredApi))
        result.requiredApi = source.requiredApi;
    if (typeof source?.reason === "string" && SAFE_CAPABILITY_TEXT.test(source.reason))
        result.reason = source.reason;
    for (const key of ["slideIndex", "textLength", "imageLength", "beforeLength", "afterLength"]) {
        const number = safeNumber(source?.[key]);
        if (number !== undefined)
            result[key] = number;
    }
    return Object.keys(result).length ? result : undefined;
}
function makeTrustedToolError(category, error, extras = {}) {
    const result = { ok: false, error, category, ...extras };
    Object.defineProperty(result, trustedToolError, { value: true, enumerable: false });
    return result;
}
function asToolError(value, fallbackCategory = "internal", options = {}) {
    const record = value && typeof value === "object" ? value : undefined;
    const trusted = Boolean(record && record[trustedToolError] === true);
    const hasExplicitCategory = isCategory(record?.category);
    const existingCategory = hasExplicitCategory ? record?.category : record?.code ? "office_host" : fallbackCategory;
    const explicitStatus = safeNumber(record?.statusCode);
    const status = safeNumber(record?.status);
    const statusCode = status ?? explicitStatus;
    const name = safeName(value instanceof Error ? value.name : record?.name);
    if (name === "AbortError")
        return makeTrustedToolError("cancellation", FIXED_MESSAGES.cancellation, { name });
    if (name === "TimeoutError")
        return makeTrustedToolError("timeout", FIXED_MESSAGES.timeout, { name });
    if (value instanceof TypeError && (options.fetchTypeError === true || fallbackCategory === "network"))
        return makeTrustedToolError("network", FIXED_MESSAGES.network, { name: "TypeError" });
    const category = statusCode !== undefined && !trusted && !hasExplicitCategory ? "provider" : existingCategory;
    const details = sanitizedDetails(record, trusted, existingCategory === "office_host");
    const extras = {
        ...(statusCode === undefined ? {} : { statusCode }),
        ...(name ? { name } : {}),
        ...(details ? { details } : {}),
        ...(safeToken(record?.requestId) ? { requestId: safeToken(record?.requestId) } : {}),
    };
    const trustedMessage = typeof record?.error === "string" ? record.error : messageOf(value);
    const error = trusted ? trustedMessage : FIXED_MESSAGES[category];
    return makeTrustedToolError(category, error, extras);
}
function stringifyToolError(value, fallbackCategory) { return JSON.stringify(asToolError(value, fallbackCategory)); }
function errorToToolError(error) { return asToolError(error, "internal", { fetchTypeError: false }); }
function categorizeError(error) { return asToolError(error, "internal", { fetchTypeError: false }).category; }
const validationError = (message, name) => { const safe = safeName(name); return makeTrustedToolError("validation", message, safe ? { name: safe } : {}); };
const timeoutError = (message) => asToolError(new DOMException(message, "TimeoutError"));
const cancellationError = (message) => asToolError(new DOMException(message, "AbortError"));
const repetitionError = (message) => makeTrustedToolError("repetition", message);
const officeHostError = (message) => makeTrustedToolError("office_host", message);
const networkError = (message) => makeTrustedToolError("network", message);
const providerError = (message, statusCode) => makeTrustedToolError("provider", message, statusCode === undefined ? {} : { statusCode });
const protocolError = (message) => makeTrustedToolError("protocol", message);
const internalError = (message) => makeTrustedToolError("internal", message);

;// ./src/toolValidation.ts


function parseToolArgs(raw) {
    if (!raw.trim())
        return {};
    try {
        const value = JSON.parse(raw);
        return value !== null && typeof value === "object" && !Array.isArray(value) ? value : null;
    }
    catch {
        return null;
    }
}
function validateToolCall(name, raw) {
    const toolIndex = toolSchemas.findIndex((tool) => tool.function.name === name);
    if (toolIndex < 0)
        return fail(`Unknown tool: ${name}`);
    const args = parseToolArgs(raw);
    if (args === null)
        return fail("Tool arguments must be a JSON object.");
    const schema = toolSchemas[toolIndex].function.parameters;
    const result = validateSchemaValue(args, schema, "");
    if (result)
        return fail(result);
    return { valid: true, toolIndex, parsedArgs: args };
}
function validateSchemaValue(value, schema, path = "") {
    const label = path || "arguments";
    const type = schema.type;
    if (type === "object") {
        if (value === null || typeof value !== "object" || Array.isArray(value))
            return `${label} must be an object.`;
        const object = value;
        const properties = (schema.properties && typeof schema.properties === "object" ? schema.properties : {});
        const required = Array.isArray(schema.required) ? schema.required.filter((item) => typeof item === "string") : [];
        for (const key of required)
            if (!(key in object))
                return `Missing or invalid required argument: ${path ? `${path}.` : ""}${key}`;
        if (schema.additionalProperties === false)
            for (const key of Object.keys(object))
                if (!(key in properties))
                    return `Unknown argument: ${path ? `${path}.` : ""}${key}`;
        for (const [key, child] of Object.entries(properties))
            if (key in object) {
                const error = validateSchemaValue(object[key], child, path ? `${path}.${key}` : key);
                if (error)
                    return error;
            }
        return null;
    }
    if (type === "array") {
        if (!Array.isArray(value))
            return `${label} must be an array.`;
        if (typeof schema.minItems === "number" && value.length < schema.minItems)
            return `${label} must contain at least ${schema.minItems} item(s).`;
        if (typeof schema.maxItems === "number" && value.length > schema.maxItems)
            return `${label} must contain at most ${schema.maxItems} item(s).`;
        const itemSchema = schema.items && typeof schema.items === "object" ? schema.items : null;
        if (itemSchema)
            for (let i = 0; i < value.length; i += 1) {
                const error = validateSchemaValue(value[i], itemSchema, `${label}[${i}]`);
                if (error)
                    return error;
            }
        return null;
    }
    if (type === "string") {
        if (typeof value !== "string")
            return `${label} must be a string argument.`;
        if (typeof schema.minLength === "number" && value.length < schema.minLength)
            return `${label} must have at least ${schema.minLength} characters.`;
    }
    else if (type === "number") {
        if (typeof value !== "number" || !Number.isFinite(value))
            return `${label} must be a finite number.`;
        if (schema.integer === true && !Number.isInteger(value))
            return `${label} must be an integer.`;
        if (schema.minimum !== undefined && value < Number(schema.minimum))
            return `${label} is below minimum ${schema.minimum}.`;
        if (schema.maximum !== undefined && value > Number(schema.maximum))
            return `${label} exceeds maximum ${schema.maximum}.`;
    }
    else if (type === "integer") {
        if (typeof value !== "number" || !Number.isFinite(value) || !Number.isInteger(value))
            return `${label} must be an integer.`;
        if (schema.minimum !== undefined && value < Number(schema.minimum))
            return `${label} is below minimum ${schema.minimum}.`;
        if (schema.maximum !== undefined && value > Number(schema.maximum))
            return `${label} exceeds maximum ${schema.maximum}.`;
    }
    else if (type === "boolean") {
        if (typeof value !== "boolean")
            return `${label} must be a boolean.`;
    }
    if (Array.isArray(schema.enum) && !schema.enum.some((candidate) => Object.is(candidate, value)))
        return `${label} must be one of: ${schema.enum.join(", ")}.`;
    return null;
}
function fail(message) { return Object.assign(validationError(message), { valid: false, category: ToolErrorCategory.Validation }); }
function validateToolArguments(rawArgs, schema, toolName) {
    if (!schema)
        return { ok: false, error: validationError(`Unknown tool: ${toolName || "unknown"}`) };
    let value;
    try {
        value = rawArgs.trim() ? JSON.parse(rawArgs) : {};
    }
    catch {
        return { ok: false, error: validationError("Invalid JSON arguments.") };
    }
    if (value === null || typeof value !== "object" || Array.isArray(value))
        return { ok: false, error: validationError("arguments must be an object.") };
    const error = validateSchemaValue(value, schema, "arguments");
    if (error)
        return { ok: false, error: validationError(error) };
    return { ok: true, parsedArgs: value };
}

;// ./src/powerPointTools.ts
/* unused harmony import specifier */ var powerPointTools_stringifyToolError;
/* unused harmony import specifier */ var powerPointTools_ToolErrorCategory;
/* unused harmony import specifier */ var powerPointTools_validateToolCall;


const MAX_READ_CHARS = 5000;
const MAX_OUTPUT_CHARS = 20000;
const NOT_SUPPORTED_MESSAGE = "Not supported by current PowerPoint Office.js API";
const TITLE_BOX = { left: 60, top: 35, width: 840, height: 70 };
const BODY_BOX = { left: 70, top: 125, width: 820, height: 330 };
/**
 * Check whether a specific Office.js requirement set / version is available.
 * At build-time we can only verify against installed types — runtime `Office.context.requirements.isSetSupported`
 * is only callable inside a live PowerPoint session.
 *
 * For testing and deterministic reset, the internal table can be cleared by calling resetRequirementCache().
 */
let _requirementOverrides = null;
function resetRequirementCache() {
    _requirementOverrides = new Map();
}
function _isSetSupported(name, minVersion) {
    // If overrides exist, check them first
    if (_requirementOverrides !== null) {
        const key = `${name}:${minVersion}`;
        if (_requirementOverrides.has(key)) {
            return _requirementOverrides.get(key) ?? false;
        }
    }
    // Fall back to runtime check (works in PowerPoint) or type-level default
    try {
        if (typeof Office !== "undefined" && Office.context?.requirements?.isSetSupported) {
            return Office.context.requirements.isSetSupported(name, minVersion);
        }
    }
    catch {
        // Not in Office runtime
    }
    // Default: assume unknown versions are unavailable, known stable ones are available
    const stableVersions = {
        PowerPointApi: ["1.0", "1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"],
    };
    const versions = stableVersions[name];
    if (!versions)
        return false;
    return versions.includes(minVersion);
}
function checkRequirement(name, minVersion) {
    const available = _isSetSupported(name, minVersion);
    const result = { name, minVersion, available };
    if (!available) {
        result.reason = `Requirement set "${name}" version "${minVersion}" is not supported by the current Office.js types or runtime.`;
    }
    return result;
}
/**
 * Override a requirement check result for testing. Pass null for available to remove an override.
 * This mutates the shared cache; resetRequirementCache() clears all overrides.
 */
function setRequirementOverride(name, minVersion, available) {
    if (_requirementOverrides === null) {
        _requirementOverrides = new Map();
    }
    const key = `${name}:${minVersion}`;
    if (available === null) {
        _requirementOverrides.delete(key);
    }
    else {
        _requirementOverrides.set(key, available);
    }
}
function getCapabilities() {
    const capAddSpeakerNotes = _buildSpeakerNotesCapability();
    return {
        read_slide_content: {
            name: "read_slide_content",
            available: true,
            requiredApi: "PowerPointApi 1.4",
        },
        delete_slide: {
            name: "delete_slide",
            available: true,
            requiredApi: "PowerPointApi 1.2",
        },
        duplicate_slide: {
            name: "duplicate_slide",
            available: false,
            requiredApi: "PowerPointApi (no duplicate API in stable types)",
            reason: "No Slide.duplicate() or SlideCollection.duplicate() API exists in the installed @types/office-js (v1.0.599). The Office.js PowerPoint API surface does not expose a native slide-duplication method — use insertSlidesFromBase64 with an exported slide as a workaround.",
        },
        reorder_slide: {
            name: "reorder_slide",
            available: true,
            requiredApi: "PowerPointApi 1.8",
        },
        add_speaker_notes: capAddSpeakerNotes,
        find_and_replace_text: {
            name: "find_and_replace_text",
            available: true,
            requiredApi: "PowerPointApi 1.4 (text read + write via TextRange.text)",
        },
        list_slide_layouts: {
            name: "list_slide_layouts",
            available: true,
            requiredApi: "PowerPointApi 1.3 (SlideLayout.name, SlideLayout.id, SlideMaster.name)",
        },
        set_slide_layout: {
            name: "set_slide_layout",
            available: true,
            requiredApi: "PowerPointApi 1.8 (Slide.applyLayout)",
        },
        format_text: {
            name: "format_text",
            available: true,
            requiredApi: "PowerPointApi 1.4 (TextRange.font bold/italic/color/size/name)",
        },
        add_image: {
            name: "add_image",
            available: false,
            requiredApi: "PowerPointApi (no addImage/addPicture API in stable types)",
            reason: "PowerPoint.ShapeCollection in the installed @types/office-js (v1.0.599) does not expose addImage() or addPicture(). Only addTextBox, addGeometricShape, addLine, addTable, and addGroup exist. Image insertion is not available through the current Office.js PowerPoint API surface. Use Slide.getImageAsBase64() for export only.",
        },
    };
}
function _buildSpeakerNotesCapability() {
    // Deterministic deployment contract: enabling this requires a deliberate
    // @types/office-js upgrade plus a reviewed implementation change.
    return {
        name: "add_speaker_notes",
        available: false,
        requiredApi: "PowerPointApi (no speakerNotes property on Slide in installed types)",
        reason: "The installed @types/office-js (v1.0.599) does not expose a speakerNotes property on PowerPoint.Slide. The Office.js PowerPoint speakerNotes API is not available in the current type definitions. The add_speaker_notes tool exists as an advertised schema returning an actionable unsupported response.",
    };
}
function replaceLiteralAll(source, findText, replaceText) {
    return source.split(findText).join(replaceText);
}
function validateHtmlColor(color) {
    return /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.test(color);
}
const powerPointTools = [
    {
        schema: {
            type: "function",
            function: {
                name: "get_presentation_overview",
                description: "Get high-level details about the active PowerPoint presentation.",
                parameters: {
                    type: "object",
                    properties: {
                        maxLength: {
                            type: "integer",
                            minimum: 1,
                            maximum: MAX_OUTPUT_CHARS,
                            description: "Maximum number of text characters to include in any returned text payload."
                        }
                    },
                    additionalProperties: false
                }
            }
        },
        execute: (args) => {
            const maxLength = validatePresentationOverviewMaxLength(args.maxLength);
            if (maxLength.valid === false)
                return Promise.resolve(stringifyToolError(validationError(maxLength.error)));
            return getPresentationOverview(maxLength.value);
        }
    },
    {
        schema: {
            type: "function",
            function: {
                name: "read_selected_slides",
                description: "Read the currently selected slides from the active presentation.",
                parameters: {
                    type: "object",
                    properties: {
                        maxLength: {
                            type: "number",
                            minimum: 1,
                            description: "Maximum number of text characters to return."
                        }
                    },
                    additionalProperties: false
                }
            }
        },
        execute: (args) => readSelectedSlides(optionalNumberArg(args, "maxLength") ?? MAX_READ_CHARS)
    },
    {
        schema: {
            type: "function",
            function: {
                name: "read_all_slides",
                description: "Read a summary of all slides in the active presentation.",
                parameters: {
                    type: "object",
                    properties: {
                        maxLength: {
                            type: "number",
                            minimum: 1,
                            description: "Maximum number of text characters to return."
                        }
                    },
                    additionalProperties: false
                }
            }
        },
        execute: (args) => readAllSlides(optionalNumberArg(args, "maxLength") ?? MAX_READ_CHARS)
    },
    {
        schema: {
            type: "function",
            function: {
                name: "create_slide",
                description: "Create a new slide at the end of the presentation.",
                parameters: {
                    type: "object",
                    properties: {
                        layoutId: { type: "string", description: "Optional slide layout ID." },
                        slideMasterId: { type: "string", description: "Optional slide master ID." }
                    },
                    additionalProperties: false
                }
            }
        },
        execute: (args) => createSlide({
            layoutId: optionalStringArg(args, "layoutId"),
            slideMasterId: optionalStringArg(args, "slideMasterId")
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "create_text_slide",
                description: "Create one new slide with a title and optional body/bullets using safe non-overlapping default layout. Prefer this over multiple add_text_box calls when drafting deck content.",
                parameters: {
                    type: "object",
                    properties: {
                        title: { type: "string", description: "Slide title." },
                        body: { type: "string", description: "Optional body paragraph text." },
                        bullets: {
                            type: "array",
                            items: { type: "string" },
                            description: "Optional bullet items. They are inserted into the body text box."
                        },
                        layoutId: { type: "string", description: "Optional slide layout ID." },
                        slideMasterId: { type: "string", description: "Optional slide master ID." }
                    },
                    required: ["title"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => createTextSlide({
            title: textArg(args, "title"),
            body: optionalStringArg(args, "body"),
            bullets: optionalStringArrayArg(args, "bullets"),
            layoutId: optionalStringArg(args, "layoutId"),
            slideMasterId: optionalStringArg(args, "slideMasterId")
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "add_text_box",
                description: "Add a text box to a specific slide.",
                parameters: {
                    type: "object",
                    properties: {
                        text: { type: "string" },
                        slideIndex: { type: "integer", minimum: 0 },
                        left: { type: "number", minimum: 0 },
                        top: { type: "number", minimum: 0 },
                        width: { type: "number", minimum: 1 },
                        height: { type: "number", minimum: 1 }
                    },
                    required: ["text"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => addTextBox({
            text: textArg(args, "text"),
            slideIndex: optionalNumberArg(args, "slideIndex"),
            shapeOptions: parseShapeOptions(args)
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "update_slide_text",
                description: "Replace the currently selected text in the presentation.",
                parameters: {
                    type: "object",
                    properties: {
                        text: { type: "string" },
                        slideIndex: { type: "integer", minimum: 0 },
                        shapeId: { type: "string" }
                    },
                    required: ["text"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => updateSlideText({
            text: textArg(args, "text"),
            slideIndex: optionalNumberArg(args, "slideIndex"),
            shapeId: optionalStringArg(args, "shapeId")
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "add_speaker_notes",
                description: "Add speaker notes to a slide (not supported in this PowerPoint Office.js snapshot).",
                parameters: {
                    type: "object",
                    properties: {
                        slideIndex: { type: "integer", minimum: 0 },
                        text: { type: "string" }
                    },
                    required: ["text"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => {
            const slideIdx = optionalNumberArg(args, "slideIndex");
            const text = textArg(args, "text");
            // Use capability-gated unsupported with actionable details
            const caps = getCapabilities();
            const snCap = caps.add_speaker_notes;
            return Promise.resolve(JSON.stringify({
                ok: false,
                action: "add_speaker_notes",
                category: ToolErrorCategory.OfficeHost,
                error: snCap.reason ?? NOT_SUPPORTED_MESSAGE,
                details: {
                    slideIndex: slideIdx ?? null,
                    textLength: text.length,
                    available: snCap.available,
                    requiredApi: snCap.requiredApi
                }
            }));
        }
    },
    {
        schema: {
            type: "function",
            function: {
                name: "add_shape",
                description: "Add a geometric shape to a specific slide.",
                parameters: {
                    type: "object",
                    properties: {
                        shapeType: { type: "string", description: "PowerPoint geometric shape name." },
                        slideIndex: { type: "integer", minimum: 0 },
                        left: { type: "number", minimum: 0 },
                        top: { type: "number", minimum: 0 },
                        width: { type: "number", minimum: 1 },
                        height: { type: "number", minimum: 1 }
                    },
                    required: ["shapeType"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => addShape({
            shapeType: textArg(args, "shapeType"),
            slideIndex: optionalNumberArg(args, "slideIndex"),
            shapeOptions: parseShapeOptions(args)
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "read_slide_content",
                description: "Read text content from shapes on a specific slide by index. Returns text from all text-bearing shapes (text boxes, rectangles with text, placeholders) on the slide.",
                parameters: {
                    type: "object",
                    properties: {
                        slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to read." },
                        maxLength: { type: "number", minimum: 1, description: "Maximum text characters to return per shape." }
                    },
                    required: ["slideIndex"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => readSlideContent({
            slideIndex: numberArg(args, "slideIndex"),
            maxLength: optionalNumberArg(args, "maxLength") ?? MAX_READ_CHARS
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "delete_slide",
                description: "Delete a slide by index. Has a built-in guard against deleting the last remaining slide — if the presentation has only one slide, returns an error instead.",
                parameters: {
                    type: "object",
                    properties: {
                        slideIndex: { type: "integer", minimum: 0, description: "Optional zero-based index of the slide to delete. Defaults to the selected slide." }
                    },
                    additionalProperties: false
                }
            }
        },
        execute: (args) => deleteSlide({ slideIndex: optionalNumberArg(args, "slideIndex") })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "duplicate_slide",
                description: "Duplicate a slide by index. NOT SUPPORTED — the current Office.js PowerPoint API surface (@types/office-js v1.0.599) has no slide duplication API. This tool exists to provide an actionable diagnostics response; use insertSlidesFromBase64 as an alternative workaround.",
                parameters: {
                    type: "object",
                    properties: {
                        slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to duplicate." }
                    },
                    required: ["slideIndex"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => duplicateSlide({ slideIndex: optionalNumberArg(args, "slideIndex") })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "reorder_slide",
                description: "Move a slide from its current position to a new position within the presentation using Slide.moveTo (PowerPointApi 1.8).",
                parameters: {
                    type: "object",
                    properties: {
                        fromIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to move." },
                        toIndex: { type: "integer", minimum: 0, description: "Zero-based target position for the slide." }
                    },
                    required: ["fromIndex", "toIndex"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => reorderSlide({
            fromIndex: numberArg(args, "fromIndex"),
            toIndex: numberArg(args, "toIndex")
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "find_and_replace_text",
                description: "Find and replace text on a slide. Searches shape text content using JavaScript-side string matching and replaces all occurrences. If shapeId is provided, only that shape is affected; otherwise searches all text-bearing shapes on the slide.",
                parameters: {
                    type: "object",
                    properties: {
                        findText: { type: "string", description: "The text string to search for." },
                        replaceText: { type: "string", description: "The replacement text string." },
                        slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide. Defaults to selected slide." },
                        shapeId: { type: "string", description: "Optional shape ID to limit the search to one shape." }
                    },
                    required: ["findText", "replaceText"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => findAndReplaceText({
            findText: textArg(args, "findText"),
            replaceText: textArg(args, "replaceText"),
            slideIndex: optionalNumberArg(args, "slideIndex"),
            shapeId: optionalStringArg(args, "shapeId")
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "list_slide_layouts",
                description: "List all slide masters and their available layouts. Returns slide master names/id and the layout names/types for each master.",
                parameters: {
                    type: "object",
                    properties: {},
                    additionalProperties: false
                }
            }
        },
        execute: () => listSlideLayouts()
    },
    {
        schema: {
            type: "function",
            function: {
                name: "set_slide_layout",
                description: "Change the layout of a specific slide by applying a layout ID. Uses Slide.applyLayout() (PowerPointApi 1.8).",
                parameters: {
                    type: "object",
                    properties: {
                        slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide to modify." },
                        layoutId: { type: "string", description: "The ID of the layout to apply (obtained from list_slide_layouts)." }
                    },
                    required: ["slideIndex", "layoutId"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => setSlideLayout({
            slideIndex: numberArg(args, "slideIndex"),
            layoutId: textArg(args, "layoutId")
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "format_text",
                description: "Format text in a shape by setting font properties (bold, italic, color, size, name). Uses TextRange.font (PowerPointApi 1.4). Only provided properties are applied.",
                parameters: {
                    type: "object",
                    properties: {
                        slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the slide." },
                        shapeId: { type: "string", description: "ID of the shape containing the text to format." },
                        font: {
                            type: "object",
                            properties: {
                                bold: { type: "boolean", description: "Bold state (true/false)." },
                                italic: { type: "boolean", description: "Italic state (true/false)." },
                                color: { type: "string", description: "HTML color code (e.g., '#FF0000' for red)." },
                                size: { type: "number", description: "Font size in points (e.g., 18)." },
                                name: { type: "string", description: "Font family name (e.g., 'Calibri')." }
                            },
                            additionalProperties: false
                        }
                    },
                    required: ["slideIndex", "shapeId"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => formatText({
            slideIndex: numberArg(args, "slideIndex"),
            shapeId: textArg(args, "shapeId"),
            font: args.font || {}
        })
    },
    {
        schema: {
            type: "function",
            function: {
                name: "add_image",
                description: "Add an image from a base64-encoded PNG or JPEG to a slide. NOT SUPPORTED — the current Office.js PowerPoint ShapeCollection API has no addImage method. This tool provides an actionable diagnostics response.",
                parameters: {
                    type: "object",
                    properties: {
                        base64Image: { type: "string", description: "Base64-encoded image data (PNG or JPEG)." },
                        slideIndex: { type: "integer", minimum: 0, description: "Zero-based index of the target slide." },
                        left: { type: "number", minimum: 0, description: "Left position in points." },
                        top: { type: "number", minimum: 0, description: "Top position in points." },
                        width: { type: "number", minimum: 1, description: "Width in points." },
                        height: { type: "number", minimum: 1, description: "Height in points." }
                    },
                    required: ["base64Image", "slideIndex"],
                    additionalProperties: false
                }
            }
        },
        execute: (args) => {
            const base64 = textArg(args, "base64Image");
            const slideIdx = numberArg(args, "slideIndex");
            const caps = getCapabilities();
            const imgCap = caps.add_image;
            return Promise.resolve(JSON.stringify({
                ok: false,
                action: "add_image",
                category: ToolErrorCategory.OfficeHost,
                error: imgCap.reason ?? NOT_SUPPORTED_MESSAGE,
                details: {
                    slideIndex: slideIdx,
                    imageLength: base64.length,
                    available: imgCap.available,
                    requiredApi: imgCap.requiredApi,
                    reason: imgCap.reason
                }
            }));
        }
    }
];
const toolSchemas = powerPointTools.map((tool) => tool.schema);
async function executePowerPointToolParsed(name, parsedArgs) {
    const tool = powerPointTools.find((entry) => entry.schema.function.name === name);
    if (!tool)
        return stringifyToolError({ ok: false, error: `Unknown tool: ${name}`, category: ToolErrorCategory.Validation });
    try {
        return normalizeSerializedToolResult(await tool.execute(parsedArgs), ToolErrorCategory.OfficeHost);
    }
    catch (error) {
        return normalizeSerializedToolResult(stringifyError(error), ToolErrorCategory.OfficeHost);
    }
}
const TOOL_ERROR_CATEGORIES = new Set(Object.values(ToolErrorCategory));
function normalizeSerializedToolResult(serialized, fallbackCategory) {
    try {
        const parsed = JSON.parse(serialized);
        if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed))
            return serialized;
        const result = parsed;
        if (result.ok !== false)
            return serialized;
        const category = typeof result.category === "string" && TOOL_ERROR_CATEGORIES.has(result.category)
            ? result.category
            : fallbackCategory;
        return JSON.stringify({ ...result, category });
    }
    catch {
        return serialized;
    }
}
async function executePowerPointTool(name, rawArgs) {
    const validation = powerPointTools_validateToolCall(name, rawArgs);
    if (!validation.valid)
        return powerPointTools_stringifyToolError(validation);
    return executePowerPointToolParsed(name, validation.parsedArgs);
}
async function getPresentationContext() {
    return getPresentationOverview();
}
function notSupportedResult(action, details) {
    return JSON.stringify({
        ok: false,
        action,
        error: NOT_SUPPORTED_MESSAGE,
        category: powerPointTools_ToolErrorCategory.OfficeHost,
        ...(details ? { details } : {})
    });
}
async function getPresentationOverview(maxLength = MAX_READ_CHARS) {
    return safelyPowerPoint(async (context) => {
        const presentation = context.presentation;
        presentation.load("title");
        const selectedSlides = presentation.getSelectedSlides();
        selectedSlides.load("items/id,index");
        const selectedTextRange = presentation.getSelectedTextRangeOrNullObject();
        selectedTextRange.load("text");
        const slideCountResult = presentation.slides.getCount();
        await context.sync();
        const selected = selectedSlides.items.map((slide) => ({
            id: slide.id,
            index: slide.index
        }));
        const text = selectedTextRange.isNullObject ? "" : selectedTextRange.text || "";
        const selectedText = truncateText(text, maxLength);
        const resolvedTitle = (presentation.title || "").trim() || "Untitled Presentation";
        return {
            ok: true,
            action: "get_presentation_overview",
            title: resolvedTitle,
            slideCount: slideCountResult.value,
            selectedSlides: selected,
            selectedCount: selected.length,
            selectedText: selectedText.text,
            selectedTextTruncated: selectedText.truncated,
            selectedTextOriginalLength: selectedText.originalLength,
            selectedTextReturnedLength: selectedText.returnedLength
        };
    });
}
async function readSelectedSlides(maxLength) {
    return safelyPowerPoint(async (context) => {
        const presentation = context.presentation;
        const selectedSlides = presentation.getSelectedSlides();
        selectedSlides.load("items/id,index");
        const selectedTextRange = presentation.getSelectedTextRangeOrNullObject();
        selectedTextRange.load("text");
        await context.sync();
        const selected = selectedSlides.items.map((slide) => ({
            id: slide.id,
            index: slide.index
        }));
        const selectedText = selectedTextRange.isNullObject ? "" : selectedTextRange.text || "";
        const truncated = truncateText(selectedText, normalizeMaxLength(maxLength));
        const summaryText = selected.length
            ? `${selectedText || `Selected ${selected.length} slide(s).`}`
            : truncated.text;
        return {
            ok: true,
            action: "read_selected_slides",
            selectedCount: selected.length,
            selectedSlides: selected,
            truncated: truncated.truncated,
            originalLength: truncated.originalLength,
            returnedLength: truncated.returnedLength,
            text: truncateText(summaryText, normalizeMaxLength(maxLength)).text
        };
    });
}
async function readAllSlides(maxLength) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id,index");
        const slideCount = slides.getCount();
        await context.sync();
        const lines = slides.items.map((slide) => `Slide ${slide.index + 1} (id:${slide.id})`);
        const summary = lines.join("\n") || "No slides.";
        const truncated = truncateText(summary, normalizeMaxLength(maxLength));
        return {
            ok: true,
            action: "read_all_slides",
            slideCount: slideCount.value,
            truncated: truncated.truncated,
            originalLength: truncated.originalLength,
            returnedLength: truncated.returnedLength,
            text: truncated.text
        };
    });
}
async function createSlide(options) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const beforeCount = slides.getCount();
        slides.load("items/id,index");
        await context.sync();
        const beforeIds = slides.items.map((slide) => slide.id);
        const addOptions = {};
        if (options.layoutId) {
            addOptions.layoutId = options.layoutId;
        }
        if (options.slideMasterId) {
            addOptions.slideMasterId = options.slideMasterId;
        }
        const hasOptions = Object.keys(addOptions).length > 0;
        if (hasOptions) {
            slides.add(addOptions);
        }
        else {
            slides.add();
        }
        await context.sync();
        slides.load("items/id,index");
        const afterCount = slides.getCount();
        await context.sync();
        const created = resolveCreatedSlide(beforeIds, slides.items.map((slide) => ({ id: slide.id, index: slide.index })));
        return {
            ok: true,
            action: "create_slide",
            beforeCount: beforeCount.value,
            afterCount: afterCount.value,
            newSlideIndex: created.index,
            newSlideId: created.id,
            layoutId: options.layoutId ?? null,
            slideMasterId: options.slideMasterId ?? null
        };
    });
}
async function createTextSlide(options) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const beforeCount = slides.getCount();
        slides.load("items/id,index");
        await context.sync();
        const beforeIds = slides.items.map((slide) => slide.id);
        const addOptions = {};
        if (options.layoutId) {
            addOptions.layoutId = options.layoutId;
        }
        if (options.slideMasterId) {
            addOptions.slideMasterId = options.slideMasterId;
        }
        if (Object.keys(addOptions).length > 0) {
            slides.add(addOptions);
        }
        else {
            slides.add();
        }
        await context.sync();
        slides.load("items/id,index");
        const afterCount = slides.getCount();
        await context.sync();
        const created = resolveCreatedSlide(beforeIds, slides.items.map((slide) => ({ id: slide.id, index: slide.index })));
        const slideIndex = created.index;
        const slide = slides.getItem(created.id);
        const titleShape = slide.shapes.addTextBox(options.title, TITLE_BOX);
        const bodyText = buildSlideBodyText(options.body, options.bullets);
        const bodyShape = bodyText ? slide.shapes.addTextBox(bodyText, BODY_BOX) : undefined;
        titleShape.load("id");
        bodyShape?.load("id");
        await context.sync();
        return {
            ok: true,
            action: "create_text_slide",
            beforeCount: beforeCount.value,
            afterCount: afterCount.value,
            newSlideIndex: slideIndex,
            titleShapeId: titleShape.id,
            bodyShapeId: bodyShape?.id ?? null,
            titleLength: options.title.length,
            bodyLength: bodyText.length
        };
    });
}
async function addTextBox(params) {
    return safelyPowerPoint(async (context) => {
        const slide = await resolveSlide(context, params.slideIndex);
        const textBox = slide.shapes.addTextBox(params.text, defaultTextBoxOptions(params.shapeOptions));
        textBox.load("id");
        await context.sync();
        return {
            ok: true,
            action: "add_text_box",
            slideId: slide.id,
            slideIndex: slide.index,
            shapeId: textBox.id,
            textLength: params.text.length
        };
    });
}
async function updateSlideText(params) {
    return safelyPowerPoint(async (context) => {
        if (params.shapeId) {
            const slide = await resolveSlide(context, params.slideIndex);
            const shape = slide.shapes.getItem(params.shapeId);
            const textFrame = shape.getTextFrameOrNullObject();
            textFrame.load("isNullObject");
            await context.sync();
            if (textFrame.isNullObject) {
                return {
                    ok: false,
                    action: "update_slide_text",
                    error: NOT_SUPPORTED_MESSAGE,
                    details: {
                        shapeId: params.shapeId
                    }
                };
            }
            const textRange = textFrame.textRange;
            textRange.load("text");
            await context.sync();
            const beforeLength = textRange.text.length;
            textRange.text = params.text;
            await context.sync();
            return {
                ok: true,
                action: "update_slide_text",
                mode: "shape",
                slideId: slide.id,
                slideIndex: slide.index,
                shapeId: params.shapeId,
                beforeLength,
                afterLength: params.text.length
            };
        }
        const selectedTextRange = context.presentation.getSelectedTextRangeOrNullObject();
        selectedTextRange.load("text");
        await context.sync();
        if (selectedTextRange.isNullObject) {
            return {
                ok: false,
                action: "update_slide_text",
                error: "No selected text to update. Provide slide and shape identifiers or select text first."
            };
        }
        const beforeLength = selectedTextRange.text.length;
        selectedTextRange.text = params.text;
        await context.sync();
        return {
            ok: true,
            action: "update_slide_text",
            mode: "selection",
            beforeLength,
            afterLength: params.text.length
        };
    });
}
async function addShape(params) {
    return safelyPowerPoint(async (context) => {
        const slide = await resolveSlide(context, params.slideIndex);
        const shape = slide.shapes.addGeometricShape(params.shapeType, params.shapeOptions);
        shape.load("id");
        await context.sync();
        return {
            ok: true,
            action: "add_shape",
            slideId: slide.id,
            slideIndex: slide.index,
            shapeId: shape.id,
            shapeType: params.shapeType
        };
    });
}
async function readSlideContent(params) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const total = slides.getCount();
        await context.sync();
        if (total.value === 0) {
            return { ok: false, action: "read_slide_content", error: "No slides in presentation." };
        }
        const safeIndex = validateExplicitSlideIndex(params.slideIndex, total.value);
        const slide = slides.getItemAt(safeIndex);
        slide.load("id,index");
        const shapes = slide.shapes;
        shapes.load("items/id");
        await context.sync();
        const shapeItems = shapes.items;
        const textShapes = [];
        const maxLen = normalizeMaxLength(params.maxLength);
        // Office.js returns a new proxy for each getTextFrameOrNullObject() call.
        // Keep and reuse the exact proxies whose properties are loaded.
        const textFrames = shapeItems.map((shape) => ({ shape, textFrame: shape.getTextFrameOrNullObject() }));
        for (const { textFrame } of textFrames)
            textFrame.load("isNullObject,hasText");
        await context.sync();
        const loadedTextRanges = textFrames
            .filter(({ textFrame }) => !textFrame.isNullObject && textFrame.hasText)
            .map(({ shape, textFrame }) => ({ shape, textRange: textFrame.textRange }));
        for (const { textRange } of loadedTextRanges)
            textRange.load("text");
        await context.sync();
        for (const { shape, textRange } of loadedTextRanges) {
            const rawText = textRange.text ?? "";
            const truncated = truncateText(rawText, maxLen);
            textShapes.push({
                shapeId: shape.id,
                text: truncated.text,
                truncated: truncated.truncated,
                originalLength: truncated.originalLength,
                returnedLength: truncated.returnedLength
            });
        }
        return {
            ok: true,
            action: "read_slide_content",
            slideIndex: safeIndex,
            slideId: slide.id,
            shapes: textShapes,
            totalShapes: shapeItems.length,
            textBearingShapes: textShapes.length
        };
    });
}
/**
 * Delete a slide by index. Uses Slide.delete() (PowerPointApi 1.2).
 * Has a last-slide guard: refuses to delete if the presentation has only one slide.
 */
async function deleteSlide(params) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const total = slides.getCount();
        await context.sync();
        if (total.value <= 1) {
            return {
                ok: false,
                action: "delete_slide",
                error: "Cannot delete the last remaining slide. A presentation must have at least one slide."
            };
        }
        const slide = await resolveSlide(context, params.slideIndex);
        const slideIndex = slide.index;
        slide.delete();
        await context.sync();
        return {
            ok: true,
            action: "delete_slide",
            slideId: slide.id,
            slideIndex,
            remainingCount: total.value - 1
        };
    });
}
/**
 * Duplicate a slide. No native duplicate API exists in @types/office-js v1.0.599.
 * Always returns capability-gated not-supported with actionable details.
 */
async function duplicateSlide(params) {
    const caps = getCapabilities();
    const dupCap = caps.duplicate_slide;
    return JSON.stringify({
        ok: false,
        action: "duplicate_slide",
        category: ToolErrorCategory.OfficeHost,
        error: dupCap.reason ?? NOT_SUPPORTED_MESSAGE,
        details: {
            slideIndex: params.slideIndex !== undefined ? Math.floor(params.slideIndex) : null,
            available: dupCap.available,
            requiredApi: dupCap.requiredApi,
            reason: dupCap.reason
        }
    });
}
/**
 * Reorder a slide using Slide.moveTo(slideIndex) — PowerPointApi 1.8.
 */
async function reorderSlide(params) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const total = slides.getCount();
        await context.sync();
        if (total.value === 0) {
            return { ok: false, action: "reorder_slide", error: "No slides in presentation." };
        }
        const fromIdx = validateExplicitSlideIndex(params.fromIndex, total.value);
        const toIdx = validateExplicitSlideIndex(params.toIndex, total.value);
        const slide = slides.getItemAt(fromIdx);
        slide.load("id,index");
        await context.sync();
        slide.moveTo(toIdx);
        await context.sync();
        return {
            ok: true,
            action: "reorder_slide",
            slideId: slide.id,
            fromIndex: fromIdx,
            toIndex: toIdx
        };
    });
}
/**
 * Find and replace text on a slide. Uses JavaScript-side string replacement on each shape's
 * TextRange.text (PowerPointApi 1.4). If shapeId is provided, only that shape is updated.
 * If findText is empty, returns an error.
 */
async function findAndReplaceText(params) {
    if (!params.findText) {
        return JSON.stringify({
            ok: false,
            action: "find_and_replace_text",
            error: "findText cannot be empty.",
            category: ToolErrorCategory.Validation
        });
    }
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const total = slides.getCount();
        await context.sync();
        if (total.value === 0) {
            return { ok: false, action: "find_and_replace_text", error: "No slides in presentation." };
        }
        const slide = await resolveSlide(context, params.slideIndex);
        slide.load("id,index");
        await context.sync();
        const shapes = slide.shapes;
        if (params.shapeId) {
            // Target a specific shape
            const shape = shapes.getItem(params.shapeId);
            const textFrame = shape.getTextFrameOrNullObject();
            textFrame.load("isNullObject,hasText");
            await context.sync();
            if (textFrame.isNullObject || !textFrame.hasText) {
                return {
                    ok: false,
                    action: "find_and_replace_text",
                    error: "Specified shape exists but has no text frame.",
                    details: { slideIndex: slide.index, shapeId: params.shapeId }
                };
            }
            const textRange = textFrame.textRange;
            textRange.load("text");
            await context.sync();
            const before = textRange.text ?? "";
            const after = replaceLiteralAll(before, params.findText, params.replaceText);
            const replaceCount = (before.split(params.findText).length - 1);
            if (replaceCount === 0) {
                return {
                    ok: true,
                    action: "find_and_replace_text",
                    slideIndex: slide.index,
                    slideId: slide.id,
                    shapeId: params.shapeId,
                    mode: "single-shape",
                    replacements: 0,
                    message: "Search text not found in the specified shape."
                };
            }
            textRange.text = after;
            await context.sync();
            return {
                ok: true,
                action: "find_and_replace_text",
                slideIndex: slide.index,
                slideId: slide.id,
                shapeId: params.shapeId,
                mode: "single-shape",
                replacements: replaceCount
            };
        }
        // Search all text-bearing shapes on the slide
        shapes.load("items/id");
        await context.sync();
        const shapeItems = shapes.items;
        // First pass: load and retain each exact TextFrame proxy.
        const loadedTextFrames = shapeItems.map((shape) => ({
            shape,
            textFrame: shape.getTextFrameOrNullObject()
        }));
        for (const { textFrame } of loadedTextFrames) {
            textFrame.load("isNullObject,hasText");
        }
        await context.sync();
        // Second pass: retain each exact TextRange proxy whose text is loaded.
        const loadedTextRanges = loadedTextFrames
            .filter(({ textFrame }) => !textFrame.isNullObject && textFrame.hasText)
            .map(({ shape, textFrame }) => ({ shape, textRange: textFrame.textRange }));
        for (const { textRange } of loadedTextRanges) {
            textRange.load("text");
        }
        await context.sync();
        let totalReplacements = 0;
        const updatedShapes = [];
        for (const { shape, textRange } of loadedTextRanges) {
            const before = textRange.text ?? "";
            const count = (before.split(params.findText).length - 1);
            if (count > 0) {
                textRange.text = replaceLiteralAll(before, params.findText, params.replaceText);
                totalReplacements += count;
                updatedShapes.push({ shapeId: shape.id, replacements: count });
            }
        }
        if (totalReplacements > 0) {
            await context.sync();
        }
        return {
            ok: true,
            action: "find_and_replace_text",
            slideIndex: slide.index,
            slideId: slide.id,
            mode: "all-shapes",
            totalReplacements,
            updatedShapes,
            shapesWithoutMatch: shapeItems.length - updatedShapes.length
        };
    });
}
/**
 * List all slide masters and their available layouts.
 * Uses SlideMasterCollection and SlideLayoutCollection (PowerPointApi 1.3).
 */
async function listSlideLayouts() {
    return safelyPowerPoint(async (context) => {
        const masters = context.presentation.slideMasters;
        masters.load("items/id,name");
        await context.sync();
        const masterEntries = [];
        for (const master of masters.items) {
            const layouts = master.layouts;
            layouts.load("items/id,name,type");
            await context.sync();
            const layoutEntries = layouts.items.map((layout) => ({
                id: layout.id,
                name: layout.name,
                type: layout.type
            }));
            masterEntries.push({
                masterId: master.id,
                masterName: master.name,
                layouts: layoutEntries
            });
        }
        return {
            ok: true,
            action: "list_slide_layouts",
            masterCount: masters.items.length,
            masters: masterEntries
        };
    });
}
/**
 * Change the layout of a specific slide using Slide.applyLayout() (PowerPointApi 1.8).
 * The layoutId must correspond to a layout obtained from list_slide_layouts.
 */
async function setSlideLayout(params) {
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const total = slides.getCount();
        await context.sync();
        if (total.value === 0) {
            return { ok: false, action: "set_slide_layout", error: "No slides in presentation." };
        }
        const safeIndex = validateExplicitSlideIndex(params.slideIndex, total.value);
        const slide = slides.getItemAt(safeIndex);
        slide.load("id,index");
        await context.sync();
        // Find the layout object by ID across all masters
        const masters = context.presentation.slideMasters;
        masters.load("items/id");
        await context.sync();
        let targetLayout = null;
        for (const master of masters.items) {
            const layouts = master.layouts;
            layouts.load("items/id");
            await context.sync();
            for (const layout of layouts.items) {
                if (layout.id === params.layoutId) {
                    targetLayout = layout;
                    break;
                }
            }
            if (targetLayout)
                break;
        }
        if (!targetLayout) {
            return {
                ok: false,
                action: "set_slide_layout",
                error: `Layout with ID "${params.layoutId}" not found. Use list_slide_layouts to enumerate available layouts.`,
                details: { slideIndex: safeIndex, layoutId: params.layoutId }
            };
        }
        slide.applyLayout(targetLayout);
        await context.sync();
        return {
            ok: true,
            action: "set_slide_layout",
            slideIndex: safeIndex,
            slideId: slide.id,
            layoutId: params.layoutId
        };
    });
}
/**
 * Format text in a shape by setting font properties on TextRange.font.
 * Uses TextRange.font (PowerPointApi 1.4) for bold, italic, color, size, name.
 * Only non-null/defined properties in the font param are applied.
 */
async function formatText(params) {
    if (typeof params.font?.color === "string" && !validateHtmlColor(params.font.color)) {
        return JSON.stringify({
            ok: false,
            action: "format_text",
            error: "font.color must be a hexadecimal HTML color in #RGB or #RRGGBB format.",
            category: ToolErrorCategory.Validation
        });
    }
    return safelyPowerPoint(async (context) => {
        const slides = context.presentation.slides;
        const total = slides.getCount();
        await context.sync();
        if (total.value === 0) {
            return { ok: false, action: "format_text", error: "No slides in presentation." };
        }
        const safeIndex = validateExplicitSlideIndex(params.slideIndex, total.value);
        const slide = slides.getItemAt(safeIndex);
        slide.load("id,index");
        await context.sync();
        const shape = slide.shapes.getItem(params.shapeId);
        const textFrame = shape.getTextFrameOrNullObject();
        textFrame.load("hasText");
        await context.sync();
        if (!textFrame.hasText) {
            return {
                ok: false,
                action: "format_text",
                error: "Specified shape has no text frame.",
                details: { slideIndex: safeIndex, shapeId: params.shapeId }
            };
        }
        const textRange = textFrame.textRange;
        const font = textRange.font;
        const fontInput = params.font || {};
        // Apply only the properties that are provided
        const applied = [];
        if (typeof fontInput.bold === "boolean") {
            font.bold = fontInput.bold;
            applied.push("bold");
        }
        if (typeof fontInput.italic === "boolean") {
            font.italic = fontInput.italic;
            applied.push("italic");
        }
        if (typeof fontInput.color === "string" && fontInput.color) {
            font.color = fontInput.color;
            applied.push("color");
        }
        if (typeof fontInput.size === "number" && fontInput.size > 0) {
            font.size = fontInput.size;
            applied.push("size");
        }
        if (typeof fontInput.name === "string" && fontInput.name) {
            font.name = fontInput.name;
            applied.push("name");
        }
        if (applied.length === 0) {
            return {
                ok: false,
                action: "format_text",
                error: "No valid font properties provided. Use bold, italic, color, size, or name.",
                category: ToolErrorCategory.Validation,
                details: { slideIndex: safeIndex, shapeId: params.shapeId }
            };
        }
        await context.sync();
        return {
            ok: true,
            action: "format_text",
            slideIndex: safeIndex,
            slideId: slide.id,
            shapeId: params.shapeId,
            applied
        };
    });
}
async function resolveSlide(context, slideIndex) {
    const slides = context.presentation.slides;
    const selectedSlides = context.presentation.getSelectedSlides();
    slides.load("items");
    selectedSlides.load("items/index");
    const total = slides.getCount();
    await context.sync();
    if (total.value === 0) {
        throw new Error("No slides available in this presentation.");
    }
    let slide;
    if (typeof slideIndex === "number") {
        const safeIndex = validateExplicitSlideIndex(slideIndex, total.value);
        slide = slides.getItemAt(safeIndex);
    }
    else {
        const firstSelectedIndex = selectedSlides.items[0]?.index;
        slide = slides.getItemAt(typeof firstSelectedIndex === "number" ? firstSelectedIndex : 0);
    }
    slide.load("id,index");
    await context.sync();
    return slide;
}
function parseShapeOptions(args) {
    const options = {};
    const left = optionalNumberArg(args, "left");
    if (left !== undefined) {
        options.left = left;
    }
    const top = optionalNumberArg(args, "top");
    if (top !== undefined) {
        options.top = top;
    }
    const width = optionalNumberArg(args, "width");
    if (width !== undefined) {
        options.width = width;
    }
    const height = optionalNumberArg(args, "height");
    if (height !== undefined) {
        options.height = height;
    }
    return options;
}
function defaultTextBoxOptions(options) {
    const hasPlacement = [options.left, options.top, options.width, options.height].some((value) => typeof value === "number");
    return hasPlacement ? options : { ...BODY_BOX };
}
function buildSlideBodyText(body, bullets) {
    const parts = [];
    const cleanBody = body?.trim();
    if (cleanBody) {
        parts.push(cleanBody);
    }
    const cleanBullets = (bullets ?? []).map((item) => item.trim()).filter(Boolean);
    if (cleanBullets.length) {
        parts.push(cleanBullets.map((item) => `• ${item}`).join("\n"));
    }
    return parts.join("\n\n");
}
function validatePresentationOverviewMaxLength(raw) {
    if (raw === undefined) {
        return { valid: true, value: MAX_READ_CHARS };
    }
    if (typeof raw !== "number" || !Number.isFinite(raw)) {
        return { valid: false, error: "maxLength must be a finite number." };
    }
    if (!Number.isInteger(raw)) {
        return { valid: false, error: "maxLength must be an integer." };
    }
    if (raw < 1) {
        return { valid: false, error: "maxLength is below minimum 1." };
    }
    if (raw > MAX_OUTPUT_CHARS) {
        return { valid: false, error: `maxLength exceeds maximum ${MAX_OUTPUT_CHARS}.` };
    }
    return { valid: true, value: raw };
}
function normalizeMaxLength(raw) {
    if (!Number.isFinite(raw) || raw < 1) {
        return MAX_READ_CHARS;
    }
    return Math.min(Math.max(Math.floor(raw), 1), MAX_OUTPUT_CHARS);
}
function truncateText(value, maxLength) {
    const original = value ?? "";
    const normalized = normalizeMaxLength(maxLength);
    const text = original.slice(0, normalized);
    const returnedLength = text.length;
    return {
        text,
        truncated: returnedLength < original.length,
        originalLength: original.length,
        returnedLength
    };
}
function textArg(args, name) {
    const value = args[name];
    if (typeof value !== "string") {
        throw new Error(`Missing or invalid string argument "${name}".`);
    }
    return value;
}
function optionalStringArg(args, name) {
    const value = args[name];
    if (value === undefined) {
        return undefined;
    }
    if (typeof value !== "string") {
        throw new Error(`Invalid string argument "${name}".`);
    }
    const trimmed = value.trim();
    return trimmed.length ? trimmed : undefined;
}
function optionalStringArrayArg(args, name) {
    const value = args[name];
    if (value === undefined) {
        return undefined;
    }
    if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) {
        throw new Error(`Invalid string array argument "${name}".`);
    }
    return value;
}
function numberArg(args, name) {
    const value = args[name];
    if (typeof value === "number" && Number.isFinite(value)) {
        return Math.floor(value);
    }
    throw new Error(`Missing or invalid number argument "${name}".`);
}
function optionalNumberArg(args, name) {
    const value = args[name];
    if (value === undefined) {
        return undefined;
    }
    if (typeof value !== "number" || !Number.isFinite(value)) {
        throw new Error(`Invalid number argument "${name}".`);
    }
    return Math.floor(value);
}
async function safelyPowerPoint(work) {
    try {
        const result = await PowerPoint.run(work);
        return normalizeSerializedToolResult(JSON.stringify(result), ToolErrorCategory.OfficeHost);
    }
    catch (error) {
        return normalizeSerializedToolResult(stringifyError(error), ToolErrorCategory.OfficeHost);
    }
}
function stringifyError(error) {
    return stringifyToolError(errorToToolError(error));
}
function validateExplicitSlideIndex(index, count) {
    if (!Number.isFinite(index) || !Number.isInteger(index) || index < 0 || index >= count) {
        const error = new Error(`slideIndex must be an integer from 0 through ${Math.max(0, count - 1)}.`);
        error.category = ToolErrorCategory.Validation;
        throw error;
    }
    return index;
}
function resolveCreatedSlide(beforeIds, afterSlides) {
    const previous = new Set(beforeIds);
    const added = afterSlides.map((slide, position) => typeof slide === "string" ? { id: slide, index: position } : slide).filter(({ id }) => !previous.has(id));
    if (added.length !== 1)
        throw new Error("Expected exactly one new slide after creation.");
    return added[0];
}

;// ./src/agent.ts



const MODEL_CALL_TIMEOUT_MS = 20_000;
const TURN_TIMEOUT_MS = 90_000;
/** A native request runs the entire server agent, not one model completion. */
const NATIVE_TURN_TIMEOUT_MS = 300_000;
const CHECKPOINT_TIMEOUT_MS = 30_000;
const MAX_RETRIES = 2;
const ABSOLUTE_MAX_ITERATIONS = 20;
const OFFICE_OPERATION_IN_FLIGHT_MESSAGE = "A PowerPoint operation may still complete. Verify the deck state before starting another request.";
const EXTERNAL_OPERATION_IN_FLIGHT_MESSAGE = "A previous Hermes operation is still settling. Wait before starting another request.";
const SYSTEM_PROMPT = `You are Hermes, an agentic AI assistant running inside the user's active Microsoft PowerPoint presentation.
Inspect before editing, make small verifiable changes, and use tool errors to correct your next action. Use the current tool schemas exactly. If native tool calls are unavailable, emit <tool_call>{"name":"tool_name","arguments":{}}</tool_call>.`;
class HermesAgent {
    constructor() {
        this.messages = [];
        this.activeGeneration = null;
        this.pendingOperations = new Map();
    }
    get history() { return [...this.messages]; }
    reset() { this.activeGeneration = null; this.messages = []; }
    trackOperation(operation, kind) {
        this.pendingOperations.set(operation, kind);
        operation.then(() => { this.pendingOperations.delete(operation); }, () => { this.pendingOperations.delete(operation); });
        return operation;
    }
    pendingOperationError() {
        if (!this.pendingOperations.size)
            return null;
        if ([...this.pendingOperations.values()].includes("office")) {
            const toolError = officeHostError(OFFICE_OPERATION_IN_FLIGHT_MESSAGE);
            return Object.assign(new Error(toolError.error), { toolError });
        }
        const toolError = internalError(EXTERNAL_OPERATION_IN_FLIGHT_MESSAGE);
        return Object.assign(new Error(toolError.error), { toolError });
    }
    async run(userRequest, settings, callbacks, userSignal, options = {}) {
        if (this.activeGeneration)
            throw new Error("A PowerPoint Hermes turn is already running.");
        const pendingError = this.pendingOperationError();
        if (pendingError)
            throw pendingError;
        const generation = makeCorrelationId();
        this.activeGeneration = generation;
        const nativeMode = isNativeHermesMode(settings);
        const bridgeSessionId = options.bridgeSessionId?.trim() || null;
        if (nativeMode && (!bridgeSessionId || !isUuid(bridgeSessionId))) {
            const toolError = validationError("A live PowerPoint bridge session is required for native Hermes mode.");
            this.activeGeneration = null;
            throw Object.assign(new Error(toolError.error), { toolError });
        }
        const turnStarted = Date.now();
        const turnBudgetMs = nativeMode ? NATIVE_TURN_TIMEOUT_MS : TURN_TIMEOUT_MS;
        let turnDeadline = turnStarted + turnBudgetMs;
        const turnController = new AbortController();
        const timeoutTurn = () => turnController.abort(new DOMException("The turn exceeded its time limit.", "TimeoutError"));
        let turnTimer = setTimeout(timeoutTurn, turnBudgetMs);
        const extendTurn = () => {
            turnDeadline = Date.now() + turnBudgetMs;
            clearTimeout(turnTimer);
            turnTimer = setTimeout(timeoutTurn, turnBudgetMs);
        };
        const checkpointController = new AbortController();
        const unlinkCheckpoint = linkAbortSignals(turnController.signal, checkpointController);
        let checkpointStopped = false;
        let checkpointFailure;
        let nativeCheckpoint;
        const unlink = linkAbortSignals(userSignal, turnController);
        const startedHistory = [...this.messages];
        let presentationContext = "";
        let checkpointFired = false;
        const current = () => this.activeGeneration === generation;
        const ensureCurrent = () => { if (!current())
            throw new DOMException("The running Hermes generation is stale.", "AbortError"); };
        const notify = (callback, ...args) => { if (current() && callback)
            callback(...args); };
        const maybeCheckpoint = async () => {
            if (checkpointFired || !callbacks.onCheckpoint || Date.now() - turnStarted < CHECKPOINT_TIMEOUT_MS)
                return "continue";
            checkpointFired = true;
            const checkpointOperation = this.trackOperation(Promise.resolve().then(() => callbacks.onCheckpoint(checkpointController.signal)), "checkpoint");
            const decision = await raceWithAbort(checkpointOperation, turnController.signal, userSignal);
            ensureCurrent();
            if (decision === "continue")
                extendTurn();
            safeNotify(callbacks.onCheckpointResult, decision);
            return decision;
        };
        const stopAtCheckpoint = (decision) => {
            if (decision !== "stop")
                return null;
            const finalContent = "Turn stopped at the 30-second checkpoint.";
            notify(callbacks.onAssistantFinal, finalContent);
            notify(callbacks.onStatus, "Stopped at checkpoint");
            return { messages: this.history, finalContent, stopReason: "checkpoint" };
        };
        // A native model call stays in flight for the whole turn: checkpoints and
        // progress must run concurrently, not only between model calls.
        const checkpointTimer = nativeMode ? setTimeout(() => {
            if (!current() || turnController.signal.aborted)
                return;
            nativeCheckpoint = maybeCheckpoint().then((decision) => {
                if (decision === "stop") {
                    checkpointStopped = true;
                    turnController.abort(new DOMException("Stopped at checkpoint.", "AbortError"));
                }
            }).catch((error) => {
                checkpointFailure = error;
                turnController.abort(new DOMException("Checkpoint failed.", "AbortError"));
            });
        }, CHECKPOINT_TIMEOUT_MS) : undefined;
        const progressTimer = nativeMode ? setInterval(() => {
            if (current())
                safeNotify(callbacks.onStatus, `Hermes is working (${Math.floor((Date.now() - turnStarted) / 1000)}s)…`);
        }, 5_000) : undefined;
        try {
            if (userSignal.aborted)
                throw abortReason(userSignal);
            const contextOperation = this.trackOperation(Promise.resolve().then(() => getPresentationContext()), "context");
            presentationContext = await raceWithAbort(contextOperation, turnController.signal, userSignal);
            ensureCurrent();
            throwIfAborted(turnController.signal, userSignal);
            const workingMessages = this.buildTurnMessages(userRequest, presentationContext, startedHistory);
            const toolCallCounts = new Map();
            const maxIterations = Math.min(Math.max(Math.floor(Number(settings.maxIterations) || 1), 1), ABSOLUTE_MAX_ITERATIONS);
            notify(callbacks.onAssistantStart);
            for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
                ensureCurrent();
                throwIfAborted(turnController.signal, userSignal);
                notify(callbacks.onStatus, `Thinking ${iteration}/${maxIterations}`);
                const beforeModel = stopAtCheckpoint(await maybeCheckpoint());
                if (beforeModel)
                    return beforeModel;
                const response = await this.chatCompletion(settings, workingMessages, turnController.signal, userSignal, turnDeadline, callbacks, nativeMode, bridgeSessionId);
                if (nativeCheckpoint)
                    await nativeCheckpoint;
                ensureCurrent();
                throwIfAborted(turnController.signal, userSignal);
                const afterModel = stopAtCheckpoint(await maybeCheckpoint());
                if (afterModel)
                    return afterModel;
                const toolCalls = response.toolCalls.length ? response.toolCalls : parseXmlToolCalls(response.content, generation);
                if (!toolCalls.length) {
                    const finalContent = removeXmlToolCalls(response.content).trim();
                    notify(callbacks.onAssistantFinal, finalContent);
                    notify(callbacks.onStatus, "Ready");
                    this.commitIfCurrent(generation, userRequest, presentationContext, finalContent, startedHistory);
                    return { messages: this.history, finalContent };
                }
                if (nativeMode) {
                    const finalContent = removeXmlToolCalls(response.content).trim();
                    notify(callbacks.onAssistantFinal, finalContent);
                    notify(callbacks.onStatus, "Ready");
                    this.commitIfCurrent(generation, userRequest, presentationContext, finalContent, startedHistory);
                    return { messages: this.history, finalContent };
                }
                workingMessages.push({ role: "assistant", content: removeXmlToolCalls(response.content).trim() || null, tool_calls: toolCalls });
                for (const toolCall of toolCalls) {
                    ensureCurrent();
                    throwIfAborted(turnController.signal, userSignal);
                    const schema = toolSchemas.find((entry) => entry.function.name === toolCall.function.name)?.function.parameters;
                    const parsedValidation = validateToolArguments(toolCall.function.arguments, schema, toolCall.function.name);
                    const signature = parsedValidation.ok ? `${toolCall.function.name}:${canonicalizeJson(parsedValidation.parsedArgs)}` : `${toolCall.function.name}:invalid:${canonicalizeJson(toolCall.function.arguments)}`;
                    const count = (toolCallCounts.get(signature) ?? 0) + 1;
                    toolCallCounts.set(signature, count);
                    const label = `${toolCall.function.name} (${count})`;
                    notify(callbacks.onActivity, { id: toolCall.id, status: "running", label });
                    const toolStarted = Date.now();
                    let result = "";
                    let toolOk = false;
                    let toolCategory;
                    try {
                        if (count > 2) {
                            result = stringifyToolError(repetitionError("Repeated identical tool call blocked; use the prior result."));
                            toolCategory = "repetition";
                        }
                        else if (!parsedValidation.ok) {
                            result = stringifyToolError(parsedValidation.error);
                            toolCategory = parsedValidation.error.category;
                        }
                        else {
                            const officePromise = this.trackOperation(Promise.resolve().then(() => executePowerPointToolParsed(toolCall.function.name, parsedValidation.parsedArgs)), "office");
                            result = await raceWithAbort(officePromise, turnController.signal, userSignal);
                            const parsedResult = safeParse(result);
                            toolOk = parsedResult?.ok !== false;
                            toolCategory = typeof parsedResult?.category === "string" ? parsedResult.category : undefined;
                        }
                    }
                    catch (error) {
                        toolCategory = classifyTurnError(error, userSignal, turnController.signal).category;
                        throw error;
                    }
                    finally {
                        safeNotify(callbacks.onToolResult, { name: toolCall.function.name, durationMs: Date.now() - toolStarted, ok: toolOk, category: toolCategory });
                    }
                    const parsedResult = safeParse(result);
                    notify(callbacks.onActivity, { id: toolCall.id, status: parsedResult?.ok === false ? "error" : "done", label });
                    workingMessages.push({ role: "tool", tool_call_id: toolCall.id, name: toolCall.function.name, content: result });
                }
            }
            const finalContent = `I reached the maximum of ${maxIterations} agent iterations and stopped to avoid unbounded presentation changes.`;
            notify(callbacks.onAssistantFinal, finalContent);
            notify(callbacks.onStatus, "Stopped at iteration limit");
            this.commitIfCurrent(generation, userRequest, presentationContext, finalContent, startedHistory);
            return { messages: this.history, finalContent, stopReason: "iteration_limit" };
        }
        catch (error) {
            if (checkpointStopped && !userSignal.aborted)
                return stopAtCheckpoint("stop");
            const toolError = classifyTurnError(checkpointFailure ?? error, userSignal, checkpointFailure ? new AbortController().signal : turnController.signal);
            if (toolError.category === "cancellation")
                throw Object.assign(new DOMException(toolError.error, "AbortError"), { toolError });
            if (toolError.category === "timeout")
                throw Object.assign(new DOMException(toolError.error, "TimeoutError"), { toolError });
            const safeError = new Error(toolError.error);
            if (toolError.name)
                safeError.name = toolError.name;
            if (toolError.statusCode !== undefined)
                Object.assign(safeError, { statusCode: toolError.statusCode });
            if (error && typeof error === "object" && error.partial === true)
                Object.assign(safeError, { partial: true });
            throw Object.assign(safeError, { toolError });
        }
        finally {
            clearTimeout(turnTimer);
            clearTimeout(checkpointTimer);
            clearInterval(progressTimer);
            checkpointController.abort();
            unlinkCheckpoint();
            unlink();
            if (this.activeGeneration === generation)
                this.activeGeneration = null;
        }
    }
    commitIfCurrent(generation, request, context, content, base) { if (this.activeGeneration === generation)
        this.messages = [...base, this.userHistoryMessage(request, context), { role: "assistant", content }]; }
    buildTurnMessages(request, context, base) { return [{ role: "system", content: SYSTEM_PROMPT }, ...base.slice(-12), this.userHistoryMessage(request, context)]; }
    userHistoryMessage(request, context) { return { role: "user", content: `Fresh presentation context:\n${context}\n\nUser request:\n${request}` }; }
    async chatCompletion(settings, messages, turnSignal, userSignal, turnDeadline, callbacks, nativeMode, bridgeSessionId) {
        const url = chatCompletionsUrl(settings.baseUrl);
        const headers = { "Content-Type": "application/json", ...(settings.requestHeaders ?? {}) };
        if (nativeMode && bridgeSessionId)
            headers["X-Hermes-Session-Id"] = bridgeSessionId;
        const body = { model: settings.model, messages, temperature: 0.2 };
        if (!nativeMode) {
            body.tools = toolSchemas;
            body.tool_choice = "auto";
        }
        const remainingTurn = turnDeadline - Date.now();
        if (remainingTurn <= 0)
            throw new DOMException("The turn exceeded its time limit.", "TimeoutError");
        const callBudgetMs = nativeMode ? remainingTurn : Math.min(MODEL_CALL_TIMEOUT_MS, remainingTurn);
        const callDeadline = Date.now() + callBudgetMs;
        // Native calls use the extendable turn signal; a second fixed timer would
        // incorrectly kill a turn after the user chose Continue.
        const call = nativeMode ? { controller: new AbortController(), cleanup: () => { } } : deadlineSignal(callBudgetMs);
        const unlink = linkAbortSignals(turnSignal, call.controller);
        let lastError;
        try {
            for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
                throwIfAborted(call.controller.signal, userSignal);
                const started = Date.now();
                try {
                    const result = await this.oneModelCall(url, headers, body, call.controller.signal, userSignal, callbacks, nativeMode);
                    safeNotify(callbacks.onModelCall, { attempt: attempt + 1, durationMs: Date.now() - started, statusCode: result.statusCode ?? 200, outcome: "success" });
                    return result;
                }
                catch (error) {
                    lastError = error;
                    const classified = classifyTurnError(error, userSignal, call.controller.signal);
                    const statusCode = statusCodeOf(error) ?? classified.statusCode;
                    safeNotify(callbacks.onModelCall, { attempt: attempt + 1, durationMs: Date.now() - started, statusCode, category: classified.category, outcome: "error" });
                    const partial = error && typeof error === "object" && error.partial === true;
                    const retryable = classified.category === "network" || statusCode === 429 || (statusCode !== undefined && statusCode >= 500 && statusCode <= 599);
                    const delay = 250 * 2 ** attempt;
                    const remaining = callDeadline - Date.now();
                    if (nativeMode || !retryable || partial || attempt >= MAX_RETRIES || remaining <= delay) {
                        if (retryable && remaining <= delay && !userSignal.aborted && !call.controller.signal.aborted)
                            throw new DOMException("The model call has insufficient time for another retry.", "TimeoutError");
                        throw error;
                    }
                    await interruptibleSleep(delay, call.controller.signal, userSignal);
                }
            }
            throw lastError;
        }
        finally {
            unlink();
            call.cleanup();
        }
    }
    async oneModelCall(url, headers, baseBody, callSignal, userSignal, callbacks, nativeMode) {
        let response;
        try {
            response = await raceWithAbort(this.trackOperation(Promise.resolve().then(() => fetch(url, { method: "POST", headers, signal: callSignal, body: JSON.stringify({ ...baseBody, stream: true }) })), "model"), callSignal, userSignal);
        }
        catch (error) {
            throw classifyFetchError(error, userSignal, callSignal);
        }
        this.observeResponse(response, callbacks, callSignal, userSignal);
        if (!response.ok) {
            const text = await readResponseBody(response, callSignal, 2_000_000, this.trackOperation.bind(this));
            const parsed = tryParseNonStreamResponse(text);
            if (parsed) {
                if (parsed.content)
                    callbacks.onAssistantDelta(parsed.content);
                return { ...parsed, statusCode: response.status, emitted: Boolean(parsed.content || parsed.toolCalls.length) };
            }
            throw providerError(`Model endpoint returned ${response.status}.`, response.status);
        }
        if (isJsonContentType(response.headers.get("content-type"))) {
            const text = await readResponseBody(response, callSignal, 2_000_000, this.trackOperation.bind(this));
            const parsed = tryParseNonStreamResponse(text);
            if (!parsed)
                throw protocolError("Model endpoint returned invalid JSON.");
            if (parsed.content)
                callbacks.onAssistantDelta(parsed.content);
            return { ...parsed, statusCode: response.status, emitted: Boolean(parsed.content || parsed.toolCalls.length) };
        }
        const emptyResponse = () => {
            // A native request may already have edited the deck. Even a zero-byte
            // response is not evidence that dispatch failed: never replay it as JSON.
            if (nativeMode)
                throw protocolError("Hermes returned no completion. Presentation changes may have occurred; review the deck before retrying.");
            return this.fallbackNonStream(url, headers, baseBody, callSignal, userSignal, callbacks);
        };
        if (!response.body)
            return emptyResponse();
        const stream = await readOpenAIStream(response.body, callbacks, callSignal, this.trackOperation.bind(this));
        if (!stream.content && stream.toolCalls.length === 0) {
            if (stream.emitted)
                throw Object.assign(new Error("The model stream ended after partial output."), { partial: true, category: "protocol" });
            return emptyResponse();
        }
        return { ...stream, statusCode: response.status };
    }
    async fallbackNonStream(url, headers, baseBody, signal, userSignal, callbacks) {
        let response;
        try {
            response = await raceWithAbort(this.trackOperation(Promise.resolve().then(() => fetch(url, { method: "POST", headers, signal, body: JSON.stringify({ ...baseBody, stream: false }) })), "model"), signal, userSignal);
        }
        catch (error) {
            throw classifyFetchError(error, userSignal, signal);
        }
        this.observeResponse(response, callbacks, signal, userSignal);
        const text = await readResponseBody(response, signal, 2_000_000, this.trackOperation.bind(this));
        const parsed = tryParseNonStreamResponse(text);
        if (!parsed)
            throw Object.assign(protocolError(`Model endpoint returned invalid non-stream response (${response.status}).`), { statusCode: response.status });
        if (parsed.content)
            callbacks.onAssistantDelta(parsed.content);
        return { ...parsed, statusCode: response.status, emitted: Boolean(parsed.content || parsed.toolCalls.length) };
    }
    observeResponse(response, callbacks, signal, userSignal) {
        try {
            callbacks.onResponseHeaders?.(response.headers, response.status);
            throwIfAborted(signal, userSignal);
        }
        catch (error) {
            // Metadata can cancel a turn before a reader exists; release that body too.
            if (response.body) {
                const body = response.body;
                void this.trackOperation(Promise.resolve().then(() => body.cancel()), "cleanup").catch(() => { });
            }
            throw error;
        }
    }
}
function canonicalizeJson(value) { const canonical = (input) => { if (Array.isArray(input))
    return `[${input.map(canonical).join(",")}]`; if (input && typeof input === "object")
    return `{${Object.keys(input).sort().map((key) => `${JSON.stringify(key)}:${canonical(input[key])}`).join(",")}}`; return JSON.stringify(input); }; if (typeof value === "string") {
    const trimmed = value.trim();
    try {
        return canonical(JSON.parse(trimmed));
    }
    catch {
        return `raw:${trimmed}`;
    }
} return canonical(value); }
function generateXmlCallId(runId, sequence) { return `xml_call_${runId}_${sequence}`; }
function makeToolCallSignature(toolCall) { return `${toolCall.function.name}:${canonicalizeJson(toolCall.function.arguments)}`; }
function tryParseNonStreamResponse(bodyText) { const parsed = safeParse(bodyText); const choices = parsed?.choices; if (!Array.isArray(choices) || !choices.length)
    return null; const message = choices[0]?.message; if (!message)
    return null; const toolCalls = []; for (const raw of message.tool_calls ?? []) {
    const item = raw;
    const fn = item.function;
    if (item.type === "function" && typeof fn?.name === "string")
        toolCalls.push({ id: typeof item.id === "string" ? item.id : `call_${toolCalls.length}`, type: "function", function: { name: fn.name, arguments: typeof fn.arguments === "string" ? fn.arguments : JSON.stringify(fn.arguments ?? {}) } });
} return { content: typeof message.content === "string" ? message.content : "", toolCalls, emitted: false }; }
function safeParse(value) { try {
    return JSON.parse(value);
}
catch {
    return null;
} }
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
function isUuid(value) { return UUID_PATTERN.test(value); }
function isJsonContentType(contentType) { return Boolean(contentType && /(?:^|;)\s*application\/json\s*(?:;|$)/i.test(contentType)); }
async function raceWithAbort(promise, signal, userSignal) { if (userSignal.aborted)
    throw abortReason(userSignal); if (signal.aborted)
    throw abortReason(signal); return new Promise((resolve, reject) => { const cleanup = () => { signal.removeEventListener("abort", onAbort); userSignal.removeEventListener("abort", onUserAbort); }; const onAbort = () => { cleanup(); reject(abortReason(signal)); }; const onUserAbort = () => { cleanup(); reject(abortReason(userSignal)); }; signal.addEventListener("abort", onAbort, { once: true }); userSignal.addEventListener("abort", onUserAbort, { once: true }); Promise.resolve(promise).then((value) => { cleanup(); resolve(value); }, (error) => { cleanup(); reject(error); }); }); }
function throwIfAborted(signal, userSignal) { if (userSignal?.aborted)
    throw abortReason(userSignal); if (signal.aborted)
    throw abortReason(signal); }
function abortReason(signal) { const reason = signal.reason; return reason instanceof DOMException ? reason : new DOMException("The running Hermes task was stopped.", "AbortError"); }
function classifyFetchError(error, userSignal, signal) { if (userSignal.aborted)
    return asToolError(new DOMException("The running Hermes task was stopped.", "AbortError")); if (signal.aborted && signal.reason?.name === "TimeoutError")
    return asToolError(new DOMException("The operation exceeded its time limit.", "TimeoutError")); return error instanceof TypeError ? asToolError(error, "network", { fetchTypeError: true }) : asToolError(error, "internal", { fetchTypeError: false }); }
function classifyTurnError(error, userSignal, signal) {
    if (userSignal.aborted)
        return asToolError(new DOMException("The running Hermes task was stopped.", "AbortError"));
    if (signal.aborted && signal.reason?.name === "TimeoutError")
        return asToolError(new DOMException("The operation exceeded its time limit.", "TimeoutError"));
    if (error && typeof error === "object" && "toolError" in error)
        return asToolError(error.toolError);
    return asToolError(error, "internal", { fetchTypeError: false });
}
function statusCodeOf(error) { return error && typeof error === "object" && typeof error.statusCode === "number" ? error.statusCode : undefined; }
function makeCorrelationId() { return `run_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`; }
function safeNotify(callback, ...args) { try {
    callback?.(...args);
}
catch { /* metadata must never change orchestration */ } }
function isBlockedRepeatedToolCall(count) { return count > 2; }
function chatCompletionsUrl(baseUrl) { const trimmed = baseUrl.trim().replace(/\/+$/, ""); const without = trimmed.replace(/\/chat\/completions$/i, ""); if (without === "http://127.0.0.1:8643/v1" || without === "http://localhost:8643/v1")
    return "/hermes/v1/chat/completions"; return trimmed.endsWith("/chat/completions") ? trimmed : `${trimmed}/chat/completions`; }
function isNativeHermesMode(settings) {
    if (settings.model.trim() !== "hermes-agent")
        return false;
    const base = settings.baseUrl.trim().replace(/\/+$/, "");
    if (base === "/hermes/v1")
        return true;
    try {
        const parsed = new URL(base);
        const loopback = parsed.hostname === "127.0.0.1" || parsed.hostname === "localhost" || parsed.hostname === "::1";
        const deskPort = ["47650", "47651", "47652", "47653", "47654"].includes(parsed.port);
        const verifiedDesk = parsed.protocol === "https:" && parsed.hostname === "localhost" && deskPort;
        const legacyLoopback = parsed.protocol === "http:" && loopback && (parsed.port === "8646" || parsed.port === "8643" || deskPort);
        return (verifiedDesk || legacyLoopback) && parsed.pathname === "/v1" && !parsed.username && !parsed.password && !parsed.search && !parsed.hash;
    }
    catch {
        return false;
    }
}
function deadlineSignal(ms) { const controller = new AbortController(); const timer = setTimeout(() => controller.abort(new DOMException("The model call exceeded its deadline.", "TimeoutError")), ms); return { controller, cleanup: () => clearTimeout(timer) }; }
function linkAbortSignals(primary, secondary) { const abort = () => secondary.abort(primary.reason ?? new DOMException("The running Hermes task was stopped.", "AbortError")); if (primary.aborted)
    abort();
else
    primary.addEventListener("abort", abort, { once: true }); return () => primary.removeEventListener("abort", abort); }
async function interruptibleSleep(ms, signal, userSignal) { await new Promise((resolve, reject) => { let settled = false; let timer; const abort = () => finish(abortReason(userSignal?.aborted ? userSignal : signal)); const cleanup = () => { clearTimeout(timer); signal.removeEventListener("abort", abort); userSignal?.removeEventListener("abort", abort); }; const finish = (error) => { if (settled)
    return; settled = true; cleanup(); error === undefined ? resolve() : reject(error); }; timer = setTimeout(() => finish(), ms); if (signal.aborted || userSignal?.aborted)
    abort();
else {
    signal.addEventListener("abort", abort, { once: true });
    userSignal?.addEventListener("abort", abort, { once: true });
} }); }
function bestEffortCancel(reader, trackOperation) {
    let cancellation = null;
    return () => {
        if (!cancellation) {
            const cancelOperation = Promise.resolve().then(() => reader.cancel());
            cancellation = trackOperation ? trackOperation(cancelOperation, "cleanup") : cancelOperation;
            cancellation.then(() => undefined, () => undefined);
        }
    };
}
function raceReaderRead(reader, signal, cancel, trackOperation) {
    if (signal.aborted) {
        cancel();
        return Promise.reject(abortReason(signal));
    }
    return new Promise((resolve, reject) => {
        let settled = false;
        const cleanup = () => signal.removeEventListener("abort", onAbort);
        const finish = (callback) => {
            if (settled)
                return;
            settled = true;
            cleanup();
            callback();
        };
        const onAbort = () => {
            cancel();
            finish(() => reject(abortReason(signal)));
        };
        signal.addEventListener("abort", onAbort, { once: true });
        const originalRead = Promise.resolve().then(() => reader.read());
        const readPromise = trackOperation ? trackOperation(originalRead, "read") : originalRead;
        readPromise.then((value) => finish(() => resolve(value)), (error) => finish(() => reject(error)));
        if (signal.aborted)
            onAbort();
    });
}
async function readResponseBody(response, signal, maxBytesOrTracker = 2_000_000, trackOperation) {
    const maxBytes = typeof maxBytesOrTracker === "number" ? maxBytesOrTracker : 2_000_000;
    const tracker = typeof maxBytesOrTracker === "function" ? maxBytesOrTracker : trackOperation;
    if (!response.body)
        return "";
    const reader = response.body.getReader();
    const cancel = bestEffortCancel(reader, tracker);
    const decoder = new TextDecoder();
    const chunks = [];
    let total = 0;
    try {
        while (true) {
            const part = await raceReaderRead(reader, signal, cancel, tracker);
            if (part.done)
                break;
            const value = part.value ?? new Uint8Array();
            total += value.byteLength;
            if (total > maxBytes)
                throw new Error("Provider response exceeded the bounded body limit.");
            chunks.push(value);
        }
        return chunks.map((chunk) => decoder.decode(chunk, { stream: true })).join("") + decoder.decode();
    }
    finally {
        cancel();
    }
}
async function readOpenAIStream(body, callbacks, signal, trackOperation) {
    const reader = body.getReader();
    const cancel = bestEffortCancel(reader, trackOperation);
    const decoder = new TextDecoder();
    const calls = new Map();
    let buffer = "";
    let content = "";
    let emitted = false;
    const protocolError = () => Object.assign(new Error("The model stream contained malformed SSE data."), { category: "protocol" });
    const consumeEvent = (event) => {
        const data = [];
        for (const rawLine of event.split(/\r?\n/)) {
            if (!rawLine || rawLine.startsWith(":"))
                continue;
            const colon = rawLine.indexOf(":");
            const field = colon < 0 ? rawLine : rawLine.slice(0, colon);
            let value = colon < 0 ? "" : rawLine.slice(colon + 1);
            if (value.startsWith(" "))
                value = value.slice(1);
            if (field === "data")
                data.push(value);
        }
        if (!data.length)
            return;
        const payload = data.join("\n");
        if (payload.trim() === "[DONE]")
            return;
        let parsed;
        try {
            parsed = JSON.parse(payload);
        }
        catch {
            throw protocolError();
        }
        const delta = parsed.choices?.[0]?.delta;
        if (!delta)
            return;
        if (delta.content) {
            content += delta.content;
            emitted = true;
            callbacks.onAssistantDelta(delta.content);
        }
        for (const partial of delta.tool_calls ?? []) {
            emitted = true;
            const index = partial.index ?? 0;
            const current = calls.get(index) ?? { id: partial.id || `call_${index}`, type: "function", function: { name: "", arguments: "" } };
            if (partial.id)
                current.id = partial.id;
            if (partial.function?.name)
                current.function.name += partial.function.name;
            if (partial.function?.arguments)
                current.function.arguments += partial.function.arguments;
            calls.set(index, current);
        }
    };
    const consumeCompleteEvents = () => {
        while (true) {
            const boundary = /\r?\n\r?\n/.exec(buffer);
            if (!boundary || boundary.index === undefined)
                return;
            const event = buffer.slice(0, boundary.index);
            buffer = buffer.slice(boundary.index + boundary[0].length);
            consumeEvent(event);
        }
    };
    try {
        while (true) {
            const part = await raceReaderRead(reader, signal, cancel, trackOperation);
            if (part.done)
                break;
            buffer += decoder.decode(part.value, { stream: true });
            consumeCompleteEvents();
        }
        buffer += decoder.decode();
        consumeCompleteEvents();
        if (buffer.trim())
            throw protocolError();
        return { content, toolCalls: [...calls.values()].filter((call) => call.function.name), emitted };
    }
    catch (error) {
        if (emitted && !(error instanceof DOMException && (error.name === "AbortError" || error.name === "TimeoutError"))) {
            throw Object.assign(error instanceof Error ? error : protocolError(), { partial: true });
        }
        throw error;
    }
    finally {
        cancel();
    }
}
function parseXmlToolCalls(content, runId) { const calls = []; const regex = /<tool_call>\s*([\s\S]*?)\s*<\/tool_call>/g; let match; let index = 0; while ((match = regex.exec(content))) {
    const parsed = safeParse(match[1]);
    if (parsed?.name)
        calls.push({ id: generateXmlCallId(runId, index++), type: "function", function: { name: parsed.name, arguments: JSON.stringify(parsed.arguments ?? {}) } });
} return calls; }
function removeXmlToolCalls(content) { return content.replace(/<tool_call>[\s\S]*?<\/tool_call>/g, ""); }

;// ./src/settings.ts
const STORAGE_KEY = "hermes.powerpoint.settings";
const DEFAULT_MODEL = "hermes-agent";
const LEGACY_MODELS = new Set(["hermes-4", ""]);
const defaultSettings = {
    baseUrl: "",
    model: DEFAULT_MODEL,
    maxIterations: 15
};
function loadSettings() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
        return { ...defaultSettings };
    }
    try {
        const parsed = JSON.parse(raw);
        const normalized = normalizeSettings(parsed);
        // Older builds stored a base URL and an API key here; neither belongs in the sidebar any more.
        if ("apiKey" in parsed || "baseUrl" in parsed || normalized.model !== parsed.model) {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(persisted(normalized)));
        }
        return normalized;
    }
    catch {
        return { ...defaultSettings };
    }
}
function saveSettings(settings) {
    const normalized = normalizeSettings(settings);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(persisted(normalized)));
    return { ...normalized, baseUrl: settings.baseUrl ?? "" };
}
function persisted(settings) {
    return { model: settings.model, maxIterations: settings.maxIterations };
}
function normalizeSettings(settings) {
    const maxIterations = Number(settings.maxIterations);
    const rawModel = (settings.model || defaultSettings.model).trim();
    return {
        baseUrl: "",
        model: LEGACY_MODELS.has(rawModel) ? DEFAULT_MODEL : rawModel,
        maxIterations: Number.isFinite(maxIterations) ? Math.min(Math.max(Math.floor(maxIterations), 1), 20) : defaultSettings.maxIterations
    };
}

;// ./src/privacy.ts
const PRIVACY_LABELS = {
    local: { label: "Local bridge", tooltip: "Requests use the loopback bridge; the configured backend may still be cloud-hosted." },
    cloud: { label: "Cloud", tooltip: "Requests are sent to a known cloud model provider." },
    external: { label: "External", tooltip: "Requests are sent to a non-loopback external endpoint." },
};
function classifyPrivacy(baseUrl) {
    try {
        const url = new URL(baseUrl);
        if (url.hostname === "127.0.0.1" || url.hostname === "localhost" || url.hostname === "::1" || url.hostname === "[::1]")
            return "local";
        if (["api.openai.com", "api.anthropic.com", "api.x.ai", "generativelanguage.googleapis.com"].includes(url.hostname.toLowerCase()))
            return "cloud";
    }
    catch {
        return "external";
    }
    return "external";
}

;// ./src/turnTracer.ts
const TRACE_KEY = "hermes.powerpoint.traces";
const MAX_TRACES = 50;
const MAX_METADATA_EVENTS = 100;
function startTrace(effectiveIterationCap) {
    return { id: `trace_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`, startedAt: Date.now(), effectiveIterationCap, modelCalls: 0, toolCalls: 0, errors: 0, modelAttempts: [], tools: [] };
}
function recordModelCall(trace, metadata, durationMs, streamedOrStatus, statusCode) {
    const event = typeof metadata === "number"
        ? { attempt: metadata + 1, durationMs: Math.max(0, Math.round(durationMs ?? 0)), statusCode: typeof streamedOrStatus === "number" ? streamedOrStatus : statusCode, outcome: "success" }
        : metadata;
    trace.modelCalls += 1;
    if (event.outcome === "error")
        trace.errors += 1;
    if (trace.modelAttempts.length < MAX_METADATA_EVENTS)
        trace.modelAttempts.push({ attempt: Math.max(1, Math.floor(event.attempt)), durationMs: Math.max(0, Math.round(event.durationMs)), ...(event.statusCode === undefined ? {} : { statusCode: event.statusCode }), ...(event.category ? { category: event.category.slice(0, 40) } : {}), outcome: event.outcome });
}
function recordToolResult(trace, name, durationMs, ok, category) {
    trace.toolCalls += 1;
    if (!ok)
        trace.errors += 1;
    if (trace.tools.length < MAX_METADATA_EVENTS)
        trace.tools.push({ name: name.slice(0, 80), durationMs: Math.max(0, Math.round(durationMs)), ok, ...(category ? { category: category.slice(0, 40) } : {}) });
}
function recordCheckpoint(trace, decision) { trace.checkpoint = decision; }
function finalizeTrace(trace) {
    trace.finishedAt = Date.now();
    try {
        const traces = loadTraces();
        traces.push(trace);
        localStorage.setItem(TRACE_KEY, JSON.stringify(traces.slice(-MAX_TRACES)));
    }
    catch {
        // Telemetry is best-effort and must never block task-pane cleanup.
    }
}
function loadTraces() { try {
    const value = JSON.parse(localStorage.getItem(TRACE_KEY) || "[]");
    return Array.isArray(value) ? value.slice(-MAX_TRACES) : [];
}
catch {
    return [];
} }
function clearTraces() { localStorage.removeItem(TRACE_KEY); }

;// ./src/checkpoint.ts
/** A one-shot, abort-aware checkpoint prompt coordinator for the task pane. */
class CheckpointController {
    constructor() {
        this.pending = null;
    }
    begin(signal) {
        if (this.pending)
            return Promise.resolve("continue");
        return new Promise((resolve) => {
            const onAbort = () => this.finish("stop");
            this.pending = { resolve, onAbort, signal };
            if (signal.aborted)
                onAbort();
            else
                signal.addEventListener("abort", onAbort, { once: true });
        });
    }
    choose(decision) {
        this.finish(decision);
    }
    cancel() {
        this.finish("stop");
    }
    get isPending() {
        return this.pending !== null;
    }
    finish(decision) {
        const pending = this.pending;
        if (!pending)
            return;
        this.pending = null;
        pending.signal.removeEventListener("abort", pending.onAbort);
        pending.resolve(decision);
    }
}

;// ./src/officeBridgeLifecycle.ts
class OfficeBridgeLifecycleController {
    constructor(createBridge, onState) {
        this.createBridge = createBridge;
        this.onState = onState;
        this.bridge = null;
        this.startPromise = null;
        this.stopPromise = null;
        this.removeVisibilityHandler = null;
        this.disposed = false;
        this.lifecycleGeneration = 0;
        this.visibilityRegistrationGeneration = 0;
        this.bridgeStateActive = false;
        this.restartWhenStartSettles = false;
        this.startWaitingForStop = false;
    }
    get sessionId() {
        if (!this.bridgeStateActive)
            return null;
        return this.bridge?.sessionId ?? null;
    }
    start() {
        if (this.disposed) {
            return Promise.reject(new Error("Bridge lifecycle is disposed"));
        }
        if (this.startPromise) {
            if (!this.startWaitingForStop && !this.bridgeStateActive)
                this.restartWhenStartSettles = true;
            return this.startPromise;
        }
        const generation = this.lifecycleGeneration;
        const stop = this.stopPromise;
        let bridge = this.bridge;
        const begin = () => {
            try {
                if (this.disposed)
                    throw new Error("Bridge lifecycle is disposed");
                if (generation !== this.lifecycleGeneration) {
                    throw new Error("Bridge lifecycle start was invalidated");
                }
                if (!this.bridge) {
                    this.createManagedBridge(generation);
                }
                bridge = this.bridge;
                if (!bridge)
                    throw new Error("Bridge unavailable");
                this.bridgeStateActive = true;
                bridge.setStateHandler?.(this.makeBridgeStateHandler(bridge, generation));
                return Promise.resolve(bridge.start());
            }
            catch (error) {
                return Promise.reject(error);
            }
        };
        this.startWaitingForStop = stop !== null;
        const underlying = stop
            ? stop.then(() => {
                this.startWaitingForStop = false;
                return begin();
            }, (error) => {
                this.startWaitingForStop = false;
                return Promise.reject(error);
            })
            : begin();
        let pending;
        pending = Promise.resolve(underlying).then((session) => {
            if (bridge && this.isCurrentBridgeRun(bridge, generation)) {
                this.onState({ connected: true, sessionId: session.sessionId });
            }
            return session;
        }, (error) => {
            if (this.isCurrentStartAttempt(bridge, generation)) {
                this.onState({
                    connected: false,
                    sessionId: null,
                    lastError: lifecycleErrorMessage(error),
                });
            }
            throw error;
        });
        this.startPromise = pending;
        void pending.then(() => this.finishStart(pending), () => this.finishStart(pending));
        return pending;
    }
    stop() {
        if (this.stopPromise)
            return this.stopPromise;
        if (!this.bridgeStateActive)
            return Promise.resolve();
        this.lifecycleGeneration += 1;
        this.bridgeStateActive = false;
        this.onState({ connected: false, sessionId: null });
        const bridge = this.bridge;
        let underlying;
        try {
            underlying = bridge ? bridge.stop() : Promise.resolve();
        }
        catch (error) {
            underlying = Promise.reject(error);
        }
        const pending = underlying;
        this.stopPromise = pending;
        void pending.then(() => this.finishStop(pending), () => this.finishStop(pending));
        return pending;
    }
    register(target, visibility) {
        if (this.disposed)
            return;
        target.addEventListener("pagehide", () => {
            void this.dispose().catch(() => {
                // Page teardown is best-effort; heartbeat TTL remains authoritative.
            });
        }, { once: true });
        if (!visibility)
            return;
        const registrationGeneration = ++this.visibilityRegistrationGeneration;
        void visibility.register((message) => {
            if (this.disposed || registrationGeneration !== this.visibilityRegistrationGeneration)
                return;
            if (message.visibilityMode === visibility.hiddenValue) {
                void this.stop().catch(() => {
                    // Stop is best-effort; session TTL cleans up an unreachable companion.
                });
            }
            else {
                if (this.startPromise)
                    this.restartWhenStartSettles = true;
                void this.start().catch(() => {
                    // start() already publishes the disconnected error state.
                });
            }
        }).then((removeHandler) => {
            if (this.disposed ||
                registrationGeneration !== this.visibilityRegistrationGeneration) {
                void Promise.resolve(removeHandler()).catch(() => {
                    // Handler cleanup during teardown is best-effort.
                });
            }
            else {
                this.removeVisibilityHandler = removeHandler;
            }
        }).catch(() => {
            // SharedRuntime visibility events are optional; heartbeat expiry is authoritative.
        });
    }
    async dispose() {
        if (this.disposed)
            return;
        this.disposed = true;
        this.visibilityRegistrationGeneration += 1;
        let stopError;
        try {
            await this.stop();
        }
        catch (error) {
            stopError = error;
        }
        const removeHandler = this.removeVisibilityHandler;
        this.removeVisibilityHandler = null;
        if (removeHandler) {
            await removeHandler();
        }
        if (stopError)
            throw stopError;
    }
    createManagedBridge(runGeneration) {
        let created = null;
        const bridgeState = this.makeBridgeStateHandlerForCreated(() => created, runGeneration);
        created = this.createBridge(bridgeState);
        this.bridge = created;
    }
    makeBridgeStateHandler(bridge, generation) {
        return (state) => {
            if (!this.isCurrentBridgeRun(bridge, generation))
                return;
            this.onState(state);
        };
    }
    makeBridgeStateHandlerForCreated(getBridge, generation) {
        return (state) => {
            const bridge = getBridge();
            if (this.disposed ||
                !this.bridgeStateActive ||
                !bridge ||
                this.bridge !== bridge ||
                this.lifecycleGeneration !== generation) {
                return;
            }
            this.onState(state);
        };
    }
    isCurrentBridgeRun(bridge, generation) {
        return !this.disposed &&
            this.bridgeStateActive &&
            this.bridge === bridge &&
            this.lifecycleGeneration === generation;
    }
    isCurrentStartAttempt(bridge, generation) {
        return !this.disposed &&
            generation === this.lifecycleGeneration &&
            (bridge === null ? this.bridge === null : this.bridge === bridge);
    }
    finishStart(pending) {
        if (this.startPromise !== pending)
            return;
        this.startWaitingForStop = false;
        this.startPromise = null;
        if (this.restartWhenStartSettles && !this.disposed && !this.bridgeStateActive) {
            this.restartWhenStartSettles = false;
            void this.start().catch(() => {
                // The visible transition owns the retry; start() publishes its state.
            });
        }
        else {
            this.restartWhenStartSettles = false;
        }
    }
    finishStop(pending) {
        if (this.stopPromise === pending)
            this.stopPromise = null;
    }
}
function lifecycleErrorMessage(error) {
    const message = error instanceof Error ? error.message : "Companion unavailable";
    if (/session.?key|document.?id|pane.?id|ticket.?id|arguments|authorization/i.test(message) ||
        /[A-Za-z0-9+/=_-]{48,}/.test(message)) {
        return "Companion unavailable";
    }
    return message || "Companion unavailable";
}

;// ./src/deskConnect.ts
// SPDX-License-Identifier: Apache-2.0
// Copyright 2026 Atlas Omnia
/**
 * Sidecar connection: find this app's HTTPS localhost bridge (bound to 127.0.0.1 by the plugin) and get the
 * user's approval in Hermes. The same file ships in the Word, Excel and PowerPoint sidebars; only the
 * DESK_APP constants differ.
 *
 * - The bridge answers only this sidebar's HTTPS origin, on its own app's port range.
 * - "Connect to Hermes" files a request; the user allows it in Hermes (/sidecar-<app>). The bridge then
 *   returns a per-sidebar grant ONCE; it is kept in this origin's localStorage so the approval survives
 *   Office and Hermes restarts. Hermes stores only a hash of it and can revoke it at any time.
 * - No API key ever reaches this page.
 */
const DESK_APP = "powerpoint";
const DESK_PORTS = [47650, 47651, 47652, 47653, 47654];
const DESK_SERVICE = `hermes-desk-${DESK_APP}`;
const DESK_COMMAND = `/sidecar-${DESK_APP}`;
const GRANT_HEADER = "X-Hermes-Desk-Grant";
const GRANT_STORAGE_KEY = `hermes.desk.${DESK_APP}.grant.v1`;
const POLL_INTERVAL_MS = 1500;
const HELLO_TIMEOUT_MS = 1500;
class ApprovalRequiredError extends Error {
    constructor(message = `Hermes approval required. Click Connect to Hermes, then allow it in Hermes with ${DESK_COMMAND}.`) {
        super(message);
        this.name = "ApprovalRequiredError";
    }
}
class DeskConnector {
    /** Browser errors are often opaque: report the observed error, not a guessed root cause. */
    get discoveryError() {
        if (!this.discoveryFailures.length)
            return "";
        return this.discoveryFailures.join("; ") +
            ". Check that the plugin is running and its local CA certificate is trusted. " +
            "A generic Load failed/Failed to fetch can mean certificate trust, connection refusal, origin rejection or a webview transport restriction; it does not identify which.";
    }
    constructor(options = {}) {
        this.memoryGrant = null;
        this.discoveryFailures = [];
        const fetchImpl = options.fetchImpl ?? globalThis.fetch;
        this.fetchImpl = fetchImpl.bind(globalThis);
        this.storage = options.storage ?? safeLocalStorage();
        this.sleep = options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
        this.ports = options.ports ?? DESK_PORTS;
        this.client = options.client ?? {};
    }
    /** The first bridge on this app's port range that identifies as this app's Sidecar bridge. */
    async discover() {
        this.discoveryFailures = [];
        for (const port of this.ports) {
            const base = `https://localhost:${port}`;
            const controller = typeof AbortController === "undefined" ? null : new AbortController();
            const timer = controller ? setTimeout(() => controller.abort(), HELLO_TIMEOUT_MS) : null;
            try {
                const response = await this.fetchImpl(`${base}/${DESK_APP}/hello`, {
                    method: "GET",
                    cache: "no-store",
                    signal: controller?.signal,
                });
                if (!response.ok) {
                    const hint = response.status === 403 || response.status === 421 ? " (origin or host rejected)" : "";
                    this.discoveryFailures.push(`${base}: HTTP ${response.status}${hint}`);
                    continue;
                }
                const body = (await response.json());
                if (body && body.service === DESK_SERVICE && body.app === DESK_APP) {
                    this.discoveryFailures = [];
                    return base;
                }
                this.discoveryFailures.push(`${base}: not this app's bridge`);
            }
            catch (error) {
                const detail = error instanceof Error ? `${error.name}: ${error.message}` : "Unknown fetch error";
                this.discoveryFailures.push(`${base}: ${detail.replace(/[\r\n\u0000-\u001f]/g, " ").slice(0, 160)}`);
            }
            finally {
                if (timer)
                    clearTimeout(timer);
            }
        }
        return null;
    }
    get grant() {
        try {
            const stored = this.storage?.getItem(GRANT_STORAGE_KEY) ?? null;
            return stored && stored.length >= 32 && stored.length <= 200 ? stored : this.memoryGrant;
        }
        catch {
            return this.memoryGrant;
        }
    }
    clearGrant() {
        this.memoryGrant = null;
        try {
            this.storage?.removeItem(GRANT_STORAGE_KEY);
        }
        catch {
            // Storage may be unavailable; the in-memory copy is already gone.
        }
    }
    saveGrant(grant) {
        this.memoryGrant = grant;
        try {
            this.storage?.setItem(GRANT_STORAGE_KEY, grant);
        }
        catch {
            // Without storage the approval lasts for this page only.
        }
    }
    /**
     * File a connect request and wait for the user's decision in Hermes. `onCode` shows the code the user
     * sees in Hermes' Allow/Deny question.
     */
    async requestApproval(base, onCode, signal) {
        const created = await this.post(base, `/${DESK_APP}/connect`, { client: this.client }, signal);
        if (created.status !== 201) {
            throw new Error(deskConnect_messageOf(created.body, "Hermes did not accept the connect request"));
        }
        const requestId = stringField(created.body, "requestId");
        const pollSecret = stringField(created.body, "pollSecret");
        const code = stringField(created.body, "code");
        const command = stringField(created.body, "command") || DESK_COMMAND;
        if (!requestId || !pollSecret || !code)
            throw new Error("Hermes returned an invalid connect request");
        onCode(code, command);
        for (;;) {
            if (signal?.aborted)
                throw new DOMException("Connect cancelled", "AbortError");
            await this.sleep(POLL_INTERVAL_MS);
            const polled = await this.post(base, `/${DESK_APP}/connect/status`, { requestId, pollSecret }, signal);
            if (polled.status !== 200)
                throw new Error(deskConnect_messageOf(polled.body, "Hermes did not answer the connect request"));
            const status = stringField(polled.body, "status");
            if (status === "pending")
                continue;
            if (status === "approved") {
                const grant = stringField(polled.body, "grant");
                if (!grant || grant.length < 32)
                    throw new Error("Hermes returned an invalid approval");
                this.saveGrant(grant);
                return "approved";
            }
            return status === "denied" ? "denied" : "expired";
        }
    }
    async post(base, path, body, signal) {
        const response = await this.fetchImpl(`${base}${path}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
            cache: "no-store",
            signal,
        });
        let parsed = null;
        try {
            parsed = await response.json();
        }
        catch {
            parsed = null;
        }
        return { status: response.status, body: parsed };
    }
}
function safeLocalStorage() {
    try {
        return typeof localStorage === "undefined" ? null : localStorage;
    }
    catch {
        return null;
    }
}
function stringField(value, key) {
    if (!value || typeof value !== "object")
        return "";
    const field = value[key];
    return typeof field === "string" ? field : "";
}
function deskConnect_messageOf(value, fallback) {
    const message = stringField(value, "error");
    return message ? message.slice(0, 300) : fallback;
}

;// ./src/officeBridge.ts

const DEFAULT_BASE_URL = "";
const PROTOCOL_VERSION = "1.0";
const HEARTBEAT_MS = 15_000;
const POLL_MS = 250;
const RECONNECT_DELAY_MS = 0;
const REGISTRATION_ATTEMPTS = 5;
const REGISTRATION_TIMEOUT_MS = 10_000;
const REGISTRATION_RETRY_DELAYS = [500, 1_000, 2_000, 4_000];
const RESULT_ATTEMPTS = 5;
const RESULT_TIMEOUT_MS = 10_000;
const RESULT_RETRY_DELAYS = [500, 1_000, 2_000, 4_000];
const MAX_RESULT_BODY_BYTES = 2 * 1024 * 1024;
const MAX_RESPONSE_ENVELOPE_BYTES = 64 * 1024;
const MAX_RESPONSE_BYTES = MAX_RESULT_BODY_BYTES + MAX_RESPONSE_ENVELOPE_BYTES;
const MAX_CACHED_RESULTS = 8;
const MAX_CACHED_RESULT_BYTES = 4 * 1024 * 1024;
const documentExecutionTails = new Map();
class BridgeHttpError extends Error {
    constructor(status, message, responseBody) {
        super(message);
        this.status = status;
        this.responseBody = responseBody;
        this.name = "BridgeHttpError";
    }
}
class BridgeTimeoutError extends Error {
    constructor() {
        super("Companion registration attempt timed out");
        this.name = "BridgeTimeoutError";
    }
}
class BridgeResultTimeoutError extends Error {
    constructor() {
        super("Companion result submission attempt timed out");
        this.name = "BridgeResultTimeoutError";
    }
}
class BridgeProtocolError extends Error {
    constructor(message) {
        super(message);
        this.name = "BridgeProtocolError";
    }
}
class BridgeCancelledError extends Error {
    constructor() {
        super("PowerPoint bridge operation was cancelled");
        this.name = "BridgeCancelledError";
    }
}
class BridgeCapacityError extends Error {
    constructor() {
        super("Companion result retry cache capacity exceeded");
        this.name = "BridgeCapacityError";
    }
}
class BridgeValidationError extends Error {
    constructor(message) {
        super(message);
        this.name = "BridgeValidationError";
    }
}
class OfficeCompanionBridge {
    constructor(deps) {
        this.deps = deps;
        this.session = null;
        this.registrationFlight = null;
        this.registrationController = null;
        this.sessionAbortController = null;
        this.completedResults = new Map();
        this.completedResultBytes = 0;
        this.reservedResultSlots = 0;
        this.reservedResultBytes = 0;
        this.heartbeatTimer = null;
        this.pollTimer = null;
        this.reconnectTimer = null;
        this.retryTimer = null;
        this.resultRetryTimer = null;
        this.cancelRetryWait = null;
        this.cancelResultRetryWait = null;
        this.pollFlight = null;
        this.heartbeatFlight = null;
        this.running = false;
        this.lifecycleGeneration = 0;
        this.sessionGeneration = 0;
        this.unregisteredLeaseKeys = new Set();
        const fetchImpl = deps.fetchImpl ?? globalThis.fetch;
        this.fetchImpl = fetchImpl.bind(globalThis);
        this.baseUrl = (deps.baseUrl ?? DEFAULT_BASE_URL).replace(/\/+$/, "");
        const setIntervalImpl = deps.setIntervalImpl ?? globalThis.setInterval;
        const clearIntervalImpl = deps.clearIntervalImpl ?? globalThis.clearInterval;
        const setTimeoutImpl = deps.setTimeoutImpl ?? globalThis.setTimeout;
        const clearTimeoutImpl = deps.clearTimeoutImpl ?? globalThis.clearTimeout;
        this.setIntervalImpl = setIntervalImpl.bind(globalThis);
        this.clearIntervalImpl = clearIntervalImpl.bind(globalThis);
        this.setTimeoutImpl = setTimeoutImpl.bind(globalThis);
        this.clearTimeoutImpl = clearTimeoutImpl.bind(globalThis);
        this.randomImpl = deps.randomImpl ?? (() => Math.random());
        this.nowImpl = deps.nowImpl ?? (() => Date.now());
        this.stateHandler = deps.onState;
        if (!/^[a-z][a-z0-9-]{0,31}$/.test(deps.hostPrefix)) {
            throw new Error("Invalid bridge host prefix");
        }
        if (!officeBridge_isUuid(deps.paneId)) {
            throw new Error("paneId must be a UUID");
        }
    }
    get sessionId() {
        return this.session?.sessionId ?? null;
    }
    get publicState() {
        return { connected: this.session !== null, sessionId: this.sessionId };
    }
    setStateHandler(handler) {
        this.stateHandler = handler;
    }
    connect() {
        if (this.session) {
            return Promise.resolve(this.publicSession(this.session));
        }
        if (this.registrationFlight) {
            return this.registrationFlight;
        }
        this.clearReconnectTimer();
        const sessionGeneration = this.reserveSessionGeneration();
        const controller = new AbortController();
        this.registrationController = controller;
        const connecting = this.connectNewSession(sessionGeneration, controller);
        this.registrationFlight = connecting;
        void connecting.then(() => {
            if (this.registrationFlight === connecting) {
                this.registrationFlight = null;
                this.registrationController = null;
            }
        }, () => {
            if (this.registrationFlight === connecting) {
                this.registrationFlight = null;
                this.registrationController = null;
            }
        });
        return connecting;
    }
    async connectNewSession(sessionGeneration, controller) {
        const documentId = await awaitWithSignal(Promise.resolve(this.deps.documentId()), controller.signal);
        if (!this.isCurrentRegistration(sessionGeneration, controller)) {
            throw new BridgeCancelledError();
        }
        if (!documentId || documentId.length > 256) {
            throw new BridgeValidationError("documentId must be a bounded non-empty string");
        }
        for (let attempt = 0; attempt < REGISTRATION_ATTEMPTS; attempt += 1) {
            if (!this.isCurrentRegistration(sessionGeneration, controller)) {
                throw new BridgeCancelledError();
            }
            if (attempt > 0) {
                const baseDelay = REGISTRATION_RETRY_DELAYS[attempt - 1];
                const jitter = Math.floor(clampRandom(this.randomImpl()) * 250);
                await this.waitForRetry(baseDelay + jitter, controller.signal);
                if (!this.isCurrentRegistration(sessionGeneration, controller)) {
                    throw new BridgeCancelledError();
                }
            }
            try {
                const lease = await this.registerAttempt({
                    documentId,
                    sessionGeneration,
                    signal: controller.signal,
                });
                if (!this.isCurrentRegistration(sessionGeneration, controller)) {
                    void this.unregisterLease(lease);
                    throw new BridgeCancelledError();
                }
                this.installSession(lease);
                return this.publicSession(lease);
            }
            catch (error) {
                if (!this.isCurrentRegistration(sessionGeneration, controller)) {
                    throw new BridgeCancelledError();
                }
                if (attempt + 1 < REGISTRATION_ATTEMPTS && isRetryableRegistrationError(error)) {
                    continue;
                }
                throw error;
            }
        }
        throw new Error("Companion registration failed");
    }
    async registerAttempt(options) {
        const attemptController = new AbortController();
        const onCampaignAbort = () => attemptController.abort();
        if (options.signal.aborted) {
            attemptController.abort();
        }
        else {
            options.signal.addEventListener("abort", onCampaignAbort, { once: true });
        }
        const operation = this.performRegistrationAttempt({
            ...options,
            signal: attemptController.signal,
        });
        try {
            return await this.awaitRegistrationAttempt(operation, options.signal, () => attemptController.abort());
        }
        catch (error) {
            if (error instanceof BridgeCancelledError || error instanceof BridgeTimeoutError) {
                let cleaned = false;
                void operation.then((lease) => {
                    if (cleaned)
                        return;
                    cleaned = true;
                    void this.unregisterLease(lease);
                }, () => {
                    // The late attempt has no validated lease to clean up.
                });
            }
            throw error;
        }
        finally {
            options.signal.removeEventListener("abort", onCampaignAbort);
        }
    }
    async performRegistrationAttempt(options) {
        const grant = this.deps.grant?.() ?? null;
        let response;
        try {
            response = await this.fetchImpl(this.url("session"), {
                method: "POST",
                headers: { "Content-Type": "application/json", ...(grant ? { [GRANT_HEADER]: grant } : {}) },
                body: JSON.stringify({
                    protocolVersion: PROTOCOL_VERSION,
                    documentId: options.documentId,
                    paneId: this.deps.paneId,
                }),
                credentials: "omit",
                redirect: "error",
                signal: options.signal,
            });
        }
        catch (error) {
            throw error;
        }
        let value;
        try {
            value = await readJsonObject(response);
        }
        catch (error) {
            throw new BridgeValidationError(error instanceof Error ? error.message : "Companion response is invalid");
        }
        if (response.status === 403 && value.approval === "required") {
            this.deps.onApprovalRequired?.();
            throw new ApprovalRequiredError();
        }
        if (response.status !== 201) {
            throw new BridgeHttpError(response.status, "Session registration failed");
        }
        const sessionId = value.sessionId;
        const sessionKey = value.sessionKey;
        const protocolVersion = value.protocolVersion;
        const expiresInSeconds = value.expiresInSeconds;
        if (!officeBridge_isUuid(sessionId) || typeof sessionKey !== "string" || sessionKey.length < 32) {
            throw new BridgeValidationError("Companion returned an invalid session identity");
        }
        if (protocolVersion !== PROTOCOL_VERSION ||
            typeof expiresInSeconds !== "number" ||
            !Number.isFinite(expiresInSeconds)) {
            throw new BridgeValidationError("Companion returned an incompatible protocol response");
        }
        return Object.freeze({
            sessionId,
            sessionKey,
            protocolVersion,
            expiresInSeconds,
            paneId: this.deps.paneId,
            documentId: options.documentId,
            sessionGeneration: options.sessionGeneration,
        });
    }
    awaitRegistrationAttempt(operation, signal, onTimeout) {
        return new Promise((resolve, reject) => {
            let settled = false;
            let timer = null;
            const cleanup = () => {
                signal.removeEventListener("abort", onAbort);
                if (timer !== null) {
                    this.clearTimeoutImpl(timer);
                    timer = null;
                }
            };
            const finish = (callback) => {
                if (settled)
                    return;
                settled = true;
                cleanup();
                callback();
            };
            const onAbort = () => finish(() => reject(new BridgeCancelledError()));
            if (signal.aborted) {
                void Promise.resolve(operation).catch(() => {
                    // Observe a fetch that ignores cancellation.
                });
                reject(new BridgeCancelledError());
                return;
            }
            signal.addEventListener("abort", onAbort, { once: true });
            timer = this.setTimeoutImpl(() => finish(() => {
                onTimeout();
                reject(new BridgeTimeoutError());
            }), REGISTRATION_TIMEOUT_MS);
            Promise.resolve(operation).then((value) => finish(() => resolve(value)), (error) => finish(() => reject(error)));
        });
    }
    waitForRetry(delayMs, signal) {
        if (signal.aborted)
            return Promise.reject(new BridgeCancelledError());
        if (this.retryTimer !== null) {
            return Promise.reject(new Error("Registration retry timer already exists"));
        }
        return new Promise((resolve, reject) => {
            let settled = false;
            let timer = null;
            const cleanup = () => {
                signal.removeEventListener("abort", onAbort);
                if (this.cancelRetryWait === cancel)
                    this.cancelRetryWait = null;
            };
            const finish = (callback) => {
                if (settled)
                    return;
                settled = true;
                if (timer !== null) {
                    this.clearTimeoutImpl(timer);
                    timer = null;
                }
                if (this.retryTimer !== null)
                    this.retryTimer = null;
                cleanup();
                callback();
            };
            const onAbort = () => finish(() => reject(new BridgeCancelledError()));
            const cancel = () => finish(() => reject(new BridgeCancelledError()));
            this.cancelRetryWait = cancel;
            signal.addEventListener("abort", onAbort, { once: true });
            timer = this.setTimeoutImpl(() => finish(resolve), delayMs);
            this.retryTimer = timer;
        });
    }
    async start() {
        const lifecycleGeneration = this.lifecycleGeneration;
        const session = await this.connect();
        const lease = this.session;
        if (lifecycleGeneration !== this.lifecycleGeneration ||
            !lease ||
            lease.sessionId !== session.sessionId ||
            !this.isCurrentLease(lease)) {
            throw new BridgeCancelledError();
        }
        if (this.running) {
            return session;
        }
        this.running = true;
        this.ensureTimers();
        return session;
    }
    heartbeatOnce() {
        const lease = this.requireSession();
        const flight = this.beginHeartbeat(lease);
        void flight.catch((error) => this.handleBackgroundError(error, lease, true));
        return flight;
    }
    beginHeartbeat(lease) {
        if (this.heartbeatFlight)
            return this.heartbeatFlight;
        const flight = this.heartbeatForLease(lease).catch((error) => {
            if (error instanceof BridgeCancelledError)
                return;
            throw error;
        });
        this.heartbeatFlight = flight;
        void flight.then(() => {
            if (this.heartbeatFlight === flight)
                this.heartbeatFlight = null;
        }, () => {
            if (this.heartbeatFlight === flight)
                this.heartbeatFlight = null;
        });
        return flight;
    }
    async heartbeatForLease(lease) {
        const signal = this.signalForLease(lease);
        if (!signal)
            return;
        const response = await awaitWithSignal(Promise.resolve(this.fetchImpl(this.url("heartbeat"), {
            method: "POST",
            headers: this.sessionHeaders(lease, true),
            body: "{}",
            credentials: "omit",
            redirect: "error",
            signal,
        })), signal);
        if (!this.isCurrentLease(lease))
            return;
        const value = await awaitWithSignal(readJsonObject(response), signal);
        if (!this.isCurrentLease(lease))
            return;
        if (response.status !== 200) {
            throw new BridgeHttpError(response.status, "Heartbeat failed");
        }
        void value;
    }
    pollOnce() {
        if (this.pollFlight)
            return this.pollFlight;
        const lease = this.requireSession();
        const flight = this.pollForLease(lease).catch((error) => {
            if (error instanceof BridgeCancelledError)
                return false;
            throw error;
        });
        this.pollFlight = flight;
        void flight.then(() => {
            if (this.pollFlight === flight)
                this.pollFlight = null;
        }, () => {
            if (this.pollFlight === flight)
                this.pollFlight = null;
        });
        void flight.catch((error) => this.handleBackgroundError(error, lease, true));
        return flight;
    }
    async pollForLease(lease) {
        const signal = this.signalForLease(lease);
        if (!signal || !this.isCurrentLease(lease))
            return false;
        const response = await awaitWithSignal(Promise.resolve(this.fetchImpl(this.url("pending"), {
            method: "GET",
            headers: this.sessionHeaders(lease, false),
            credentials: "omit",
            redirect: "error",
            signal,
        })), signal);
        if (!this.isCurrentLease(lease))
            return false;
        if (response.status === 204) {
            const cached = this.findCachedResultForLease(lease);
            return cached ? this.flushCompletedResult(lease, cached) : false;
        }
        const value = await awaitWithSignal(readJsonObject(response), signal);
        if (!this.isCurrentLease(lease))
            return false;
        if (response.status !== 200) {
            throw new BridgeHttpError(response.status, "Pending poll failed", value);
        }
        const ticketId = value.ticketId;
        const requestId = value.requestId;
        const tool = value.tool;
        const argumentsValue = value.arguments;
        if (!officeBridge_isUuid(ticketId))
            throw new BridgeValidationError("Companion ticketId must be a UUID");
        if (!officeBridge_isUuid(requestId))
            throw new BridgeValidationError("Companion requestId must be a UUID");
        if (typeof tool !== "string" || !tool || tool.length > 256) {
            throw new BridgeValidationError("Companion tool name is invalid");
        }
        if (!isRecord(argumentsValue)) {
            throw new BridgeValidationError("Companion arguments must be an object");
        }
        if (!this.isCurrentLease(lease))
            return false;
        const cached = this.completedResults.get(requestId);
        if (cached) {
            if (!this.matchesCachedRequest(cached, lease, tool, argumentsValue)) {
                throw new BridgeValidationError("Companion request conflicts with a retained result");
            }
            const updated = this.reattachCachedResult(cached, lease, ticketId);
            return this.flushCompletedResult(lease, updated);
        }
        const cachedResult = await this.executeAndCacheResult(lease, signal, ticketId, requestId, tool, argumentsValue);
        if (!cachedResult || !this.isCurrentLease(lease))
            return false;
        return this.flushCompletedResult(lease, cachedResult);
    }
    async executeAndCacheResult(lease, signal, ticketId, requestId, tool, argumentsValue) {
        let cachedAfterWait = null;
        let reservation = null;
        let result;
        try {
            reservation = this.reserveResultCapacity();
            result = await this.withDocumentExecution(lease.documentId, signal, async () => {
                const existing = this.completedResults.get(requestId);
                if (existing) {
                    if (!this.matchesCachedRequest(existing, lease, tool, argumentsValue)) {
                        throw new BridgeValidationError("Companion request conflicts with a retained result");
                    }
                    cachedAfterWait = existing;
                    return undefined;
                }
                return this.runToolExecution(lease, signal, tool, argumentsValue);
            });
            if (cachedAfterWait !== null) {
                this.releaseResultCapacity(reservation);
                reservation = null;
                return this.reattachCachedResult(cachedAfterWait, lease, ticketId);
            }
            const cached = this.cacheCompletedResult({
                ticketId,
                requestId,
                sessionId: lease.sessionId,
                sessionGeneration: lease.sessionGeneration,
                documentId: lease.documentId,
                tool,
                argumentsValue,
            }, result, reservation);
            reservation = null;
            return cached;
        }
        catch (error) {
            if (reservation !== null)
                this.releaseResultCapacity(reservation);
            if (error instanceof BridgeCancelledError && !this.isCurrentLease(lease))
                return null;
            throw error;
        }
    }
    async runToolExecution(lease, signal, tool, argumentsValue) {
        if (!this.isCurrentLease(lease))
            throw new BridgeCancelledError();
        let execution;
        try {
            execution = Promise.resolve(this.deps.executeTool(tool, argumentsValue));
        }
        catch (error) {
            return executionFailureResult(error);
        }
        try {
            const raw = await awaitWithSignal(execution, signal);
            return parseToolResult(raw);
        }
        catch (error) {
            if (!this.isCurrentLease(lease)) {
                const settled = await settleOperation(execution);
                return settled.ok ? parseToolResult(settled.value) : executionFailureResult(settled.error);
            }
            return executionFailureResult(error);
        }
    }
    findCachedResultForLease(lease) {
        for (const cached of this.completedResults.values()) {
            if (cached.documentId === lease.documentId &&
                cached.sessionId === lease.sessionId &&
                cached.sessionGeneration === lease.sessionGeneration) {
                return cached;
            }
        }
        return null;
    }
    matchesCachedRequest(cached, lease, tool, argumentsValue) {
        return cached.documentId === lease.documentId &&
            cached.tool === tool &&
            cached.argumentsFingerprint === canonicalJson(argumentsValue);
    }
    reattachCachedResult(cached, lease, ticketId) {
        const serialized = serializeResultBody(ticketId, cached.result);
        const updated = {
            ...cached,
            ticketId,
            sessionId: lease.sessionId,
            sessionGeneration: lease.sessionGeneration,
            result: serialized.result,
            serializedBody: serialized.serializedBody,
            byteLength: serialized.byteLength,
        };
        this.replaceCachedResult(updated);
        return updated;
    }
    async withDocumentExecution(documentId, signal, operation) {
        const previous = documentExecutionTails.get(documentId) ?? Promise.resolve();
        let release;
        const current = new Promise((resolve) => { release = resolve; });
        const chain = previous.then(() => current);
        documentExecutionTails.set(documentId, chain);
        void chain.then(() => {
            if (documentExecutionTails.get(documentId) === chain) {
                documentExecutionTails.delete(documentId);
            }
        });
        try {
            await awaitWithSignal(previous, signal);
            return await operation();
        }
        finally {
            release();
        }
    }
    reserveResultCapacity() {
        if (this.completedResults.size + this.reservedResultSlots + 1 > MAX_CACHED_RESULTS ||
            this.completedResultBytes + this.reservedResultBytes + MAX_RESULT_BODY_BYTES > MAX_CACHED_RESULT_BYTES) {
            throw new BridgeCapacityError();
        }
        this.reservedResultSlots += 1;
        this.reservedResultBytes += MAX_RESULT_BODY_BYTES;
        return { state: "reserved" };
    }
    releaseResultCapacity(reservation) {
        if (reservation.state !== "reserved")
            return;
        this.reservedResultSlots -= 1;
        this.reservedResultBytes -= MAX_RESULT_BODY_BYTES;
        reservation.state = "released";
    }
    consumeResultCapacity(reservation) {
        if (reservation.state !== "reserved") {
            throw new Error("Companion result capacity reservation is not active");
        }
        this.reservedResultSlots -= 1;
        this.reservedResultBytes -= MAX_RESULT_BODY_BYTES;
        reservation.state = "consumed";
    }
    cacheCompletedResult(identityOrTicketId, result, reservation) {
        const identity = typeof identityOrTicketId === "string"
            ? {
                ticketId: identityOrTicketId,
                requestId: identityOrTicketId,
                sessionId: "",
                sessionGeneration: 0,
                documentId: "",
                tool: "",
                argumentsValue: {},
            }
            : identityOrTicketId;
        const existing = this.completedResults.get(identity.requestId);
        if (existing && !reservation)
            return existing;
        const serialized = serializeResultBody(identity.ticketId, result);
        const cached = {
            requestId: identity.requestId,
            ticketId: identity.ticketId,
            originSessionId: identity.sessionId,
            originSessionGeneration: identity.sessionGeneration,
            documentId: identity.documentId,
            sessionId: identity.sessionId,
            sessionGeneration: identity.sessionGeneration,
            tool: identity.tool,
            argumentsFingerprint: canonicalJson(identity.argumentsValue),
            result: serialized.result,
            serializedBody: serialized.serializedBody,
            byteLength: serialized.byteLength,
            resultAttempts: 0,
            retryNotBefore: null,
        };
        this.replaceCachedResult(cached, reservation);
        return cached;
    }
    replaceCachedResult(cached, reservation) {
        const existing = this.completedResults.get(cached.requestId);
        const nextCount = this.completedResults.size + (existing ? 0 : 1);
        const nextBytes = this.completedResultBytes - (existing?.byteLength ?? 0) + cached.byteLength;
        if (nextCount > MAX_CACHED_RESULTS || nextBytes > MAX_CACHED_RESULT_BYTES) {
            throw new BridgeCapacityError();
        }
        if (reservation)
            this.consumeResultCapacity(reservation);
        this.completedResults.set(cached.requestId, cached);
        this.completedResultBytes = nextBytes;
    }
    deleteCompletedResult(requestId) {
        const cached = this.completedResults.get(requestId);
        if (!cached)
            return;
        this.completedResultBytes -= cached.byteLength;
        this.completedResults.delete(requestId);
    }
    async flushCompletedResult(lease, cachedResult) {
        if (!this.isCurrentLease(lease))
            return false;
        if (cachedResult.resultAttempts >= RESULT_ATTEMPTS)
            return false;
        while (cachedResult.resultAttempts < RESULT_ATTEMPTS) {
            if (!this.isCurrentLease(lease))
                return false;
            if (cachedResult.retryNotBefore !== null) {
                const remaining = Math.max(0, cachedResult.retryNotBefore - this.nowImpl());
                if (remaining > 0) {
                    try {
                        await this.waitForResultRetry(remaining, this.signalForLease(lease));
                    }
                    catch (error) {
                        if (!this.isCurrentLease(lease))
                            return false;
                        throw error;
                    }
                }
                cachedResult.retryNotBefore = null;
            }
            if (!this.isCurrentLease(lease))
                return false;
            cachedResult.resultAttempts += 1;
            try {
                await this.submitResult(lease, cachedResult.serializedBody, cachedResult.ticketId);
            }
            catch (error) {
                if (!this.isCurrentLease(lease))
                    return false;
                if (isTerminalResultRejection(error)) {
                    if (this.completedResults.get(cachedResult.requestId) === cachedResult) {
                        this.deleteCompletedResult(cachedResult.requestId);
                        return true;
                    }
                    return false;
                }
                if (!isRetryableResultError(error) || cachedResult.resultAttempts >= RESULT_ATTEMPTS) {
                    throw error;
                }
                const baseDelay = RESULT_RETRY_DELAYS[cachedResult.resultAttempts - 1];
                const jitter = Math.floor(clampRandom(this.randomImpl()) * 250);
                cachedResult.retryNotBefore = this.nowImpl() + baseDelay + jitter;
                try {
                    await this.waitForResultRetry(baseDelay + jitter, this.signalForLease(lease));
                }
                catch (waitError) {
                    if (!this.isCurrentLease(lease))
                        return false;
                    throw waitError;
                }
                if (!this.isCurrentLease(lease))
                    return false;
                cachedResult.retryNotBefore = null;
                continue;
            }
            if (!this.isCurrentLease(lease))
                return false;
            if (this.completedResults.get(cachedResult.requestId) !== cachedResult)
                return false;
            this.deleteCompletedResult(cachedResult.requestId);
            return true;
        }
        return false;
    }
    waitForResultRetry(delayMs, signal) {
        if (signal.aborted)
            return Promise.reject(new BridgeCancelledError());
        if (this.resultRetryTimer !== null) {
            return Promise.reject(new Error("Result retry timer already exists"));
        }
        return new Promise((resolve, reject) => {
            let settled = false;
            let timer = null;
            const cleanup = () => {
                signal.removeEventListener("abort", onAbort);
                if (this.cancelResultRetryWait === cancel)
                    this.cancelResultRetryWait = null;
            };
            const finish = (callback) => {
                if (settled)
                    return;
                settled = true;
                if (timer !== null) {
                    this.clearTimeoutImpl(timer);
                    timer = null;
                }
                if (this.resultRetryTimer !== null)
                    this.resultRetryTimer = null;
                cleanup();
                callback();
            };
            const onAbort = () => finish(() => reject(new BridgeCancelledError()));
            const cancel = () => finish(() => reject(new BridgeCancelledError()));
            this.cancelResultRetryWait = cancel;
            signal.addEventListener("abort", onAbort, { once: true });
            timer = this.setTimeoutImpl(() => finish(resolve), delayMs);
            this.resultRetryTimer = timer;
        });
    }
    async submitResult(lease, serializedBody, ticketId) {
        const leaseSignal = this.signalForLease(lease);
        if (!leaseSignal)
            return;
        const attemptController = new AbortController();
        const operation = (async () => {
            const response = await this.fetchImpl(this.url("result"), {
                method: "POST",
                headers: this.sessionHeaders(lease, true),
                body: serializedBody,
                credentials: "omit",
                redirect: "error",
                signal: attemptController.signal,
            });
            if (!this.isCurrentLease(lease))
                return;
            const value = await awaitWithSignal(readJsonObject(response), attemptController.signal);
            if (!this.isCurrentLease(lease))
                return;
            if (response.status !== 200) {
                throw new BridgeHttpError(response.status, "Result submission failed", value);
            }
            if (value.ok !== true || value.ticketId !== ticketId) {
                throw new BridgeProtocolError("Companion result acknowledgement was invalid");
            }
        })();
        await this.awaitResultAttempt(operation, leaseSignal, () => attemptController.abort());
    }
    awaitResultAttempt(operation, signal, onTimeout) {
        return new Promise((resolve, reject) => {
            let settled = false;
            let timer = null;
            const cleanup = () => {
                signal.removeEventListener("abort", onAbort);
                if (timer !== null) {
                    this.clearTimeoutImpl(timer);
                    timer = null;
                }
            };
            const finish = (callback) => {
                if (settled)
                    return;
                settled = true;
                cleanup();
                callback();
            };
            const onAbort = () => finish(() => {
                onTimeout();
                reject(new BridgeCancelledError());
            });
            if (signal.aborted) {
                void Promise.resolve(operation).catch(() => undefined);
                onTimeout();
                reject(new BridgeCancelledError());
                return;
            }
            signal.addEventListener("abort", onAbort, { once: true });
            timer = this.setTimeoutImpl(() => finish(() => {
                onTimeout();
                reject(new BridgeResultTimeoutError());
            }), RESULT_TIMEOUT_MS);
            Promise.resolve(operation).then((value) => finish(() => resolve(value)), (error) => finish(() => reject(error)));
        });
    }
    async stop() {
        this.lifecycleGeneration += 1;
        this.running = false;
        this.clearReconnectTimer();
        this.cancelRetryWait?.();
        this.cancelRetryWait = null;
        this.cancelResultRetryWait?.();
        this.cancelResultRetryWait = null;
        const registrationController = this.registrationController;
        registrationController?.abort();
        this.registrationController = null;
        this.registrationFlight = null;
        this.sessionAbortController?.abort();
        this.clearTimers();
        this.pollFlight = null;
        this.heartbeatFlight = null;
        const session = this.session;
        this.session = null;
        this.sessionGeneration += 1;
        this.sessionAbortController = null;
        this.emitState();
        if (session) {
            await this.unregisterLease(session);
        }
    }
    async dispose() {
        await this.stop();
    }
    ensureTimers() {
        if (this.heartbeatTimer === null) {
            this.heartbeatTimer = this.setIntervalImpl(() => {
                const lease = this.session;
                if (!this.running || !lease || !this.isCurrentLease(lease))
                    return;
                const flight = this.beginHeartbeat(lease);
                void flight.catch((error) => this.handleBackgroundError(error, lease, true));
            }, HEARTBEAT_MS);
        }
        if (this.pollTimer === null) {
            this.pollTimer = this.setIntervalImpl(() => {
                const lease = this.session;
                if (!this.running || !lease || !this.isCurrentLease(lease))
                    return;
                const flight = this.pollOnce();
                void flight.catch((error) => this.handleBackgroundError(error, lease, true));
            }, POLL_MS);
        }
    }
    clearTimers() {
        if (this.heartbeatTimer !== null) {
            this.clearIntervalImpl(this.heartbeatTimer);
            this.heartbeatTimer = null;
        }
        if (this.pollTimer !== null) {
            this.clearIntervalImpl(this.pollTimer);
            this.pollTimer = null;
        }
    }
    handleBackgroundError(error, lease, reconnectAllowed) {
        if (!this.isCurrentLease(lease) || error instanceof BridgeCancelledError)
            return;
        if (reconnectAllowed && this.running && isRetryableSessionError(error)) {
            this.invalidateSession(lease);
            this.scheduleReconnect();
            return;
        }
        this.stateHandler?.({
            connected: false,
            sessionId: this.session?.sessionId ?? null,
            lastError: backgroundErrorMessage(error),
        });
    }
    invalidateSession(lease) {
        if (!this.isCurrentLease(lease))
            return;
        this.sessionAbortController?.abort();
        this.sessionAbortController = null;
        this.session = null;
        this.sessionGeneration += 1;
        this.pollFlight = null;
        this.heartbeatFlight = null;
        this.emitState();
    }
    scheduleReconnect() {
        if (!this.running || this.session || this.registrationFlight || this.reconnectTimer !== null) {
            return;
        }
        const lifecycleGeneration = this.lifecycleGeneration;
        this.reconnectTimer = this.setTimeoutImpl(() => {
            this.reconnectTimer = null;
            if (!this.running ||
                lifecycleGeneration !== this.lifecycleGeneration ||
                this.session ||
                this.registrationFlight) {
                return;
            }
            const reconnect = this.connect();
            void reconnect.catch((error) => {
                if (lifecycleGeneration !== this.lifecycleGeneration || !this.running)
                    return;
                this.stateHandler?.({
                    connected: false,
                    sessionId: null,
                    lastError: backgroundErrorMessage(error),
                });
            });
        }, RECONNECT_DELAY_MS);
    }
    clearReconnectTimer() {
        if (this.reconnectTimer === null)
            return;
        this.clearTimeoutImpl(this.reconnectTimer);
        this.reconnectTimer = null;
    }
    reserveSessionGeneration() {
        this.sessionGeneration += 1;
        return this.sessionGeneration;
    }
    isCurrentRegistration(sessionGeneration, controller) {
        return this.sessionGeneration === sessionGeneration &&
            this.registrationController === controller &&
            !controller.signal.aborted;
    }
    isCurrentLease(lease) {
        return this.session === lease &&
            this.session?.sessionId === lease.sessionId &&
            this.sessionGeneration === lease.sessionGeneration;
    }
    installSession(lease) {
        if (!this.isCurrentRegistration(lease.sessionGeneration, this.registrationController)) {
            throw new BridgeCancelledError();
        }
        this.sessionAbortController = new AbortController();
        this.session = lease;
        this.emitState();
    }
    signalForLease(lease) {
        if (!this.isCurrentLease(lease) || !this.sessionAbortController)
            return null;
        return this.sessionAbortController.signal;
    }
    requireSession() {
        if (!this.session)
            throw new Error("PowerPoint bridge session is not connected");
        return this.session;
    }
    chatHeaders() {
        return this.session ? this.sessionHeaders(this.session, false) : null;
    }
    sessionHeaders(session, includeJson) {
        return {
            ...(includeJson ? { "Content-Type": "application/json" } : {}),
            "X-Hermes-Session-Key": session.sessionKey,
            "X-Hermes-Session-Id": session.sessionId,
            "X-Hermes-Pane-Id": session.paneId,
        };
    }
    async unregisterLease(session) {
        const key = `${session.sessionId}\u0000${session.sessionKey}`;
        if (this.unregisteredLeaseKeys.has(key))
            return;
        this.unregisteredLeaseKeys.add(key);
        try {
            await this.fetchImpl(this.url("session"), {
                method: "DELETE",
                headers: this.sessionHeaders(session, false),
                keepalive: true,
                credentials: "omit",
                redirect: "error",
            });
        }
        catch {
            // Heartbeat TTL remains the authoritative cleanup path.
        }
    }
    url(route) {
        return `${this.baseUrl}/${this.deps.hostPrefix}/${route}`;
    }
    publicSession(session) {
        return {
            sessionId: session.sessionId,
            protocolVersion: session.protocolVersion,
            expiresInSeconds: session.expiresInSeconds,
        };
    }
    emitState() {
        this.stateHandler?.(this.publicState);
    }
}
async function settleOperation(operation) {
    return Promise.resolve(operation).then((value) => ({ ok: true, value }), (error) => ({ ok: false, error }));
}
function parseToolResult(raw) {
    try {
        return JSON.parse(raw);
    }
    catch {
        return raw;
    }
}
function executionFailureResult(error) {
    const message = error instanceof Error ? error.message : "Office tool execution failed";
    return {
        ok: false,
        category: "Execution",
        error: sanitizeExecutionError(message),
    };
}
function sanitizeExecutionError(message) {
    if (/session.?key|document.?id|pane.?id|ticket.?id|authorization/i.test(message) ||
        /[A-Za-z0-9+/=_-]{48,}/.test(message)) {
        return "Office tool execution failed";
    }
    return message.slice(0, 512) || "Office tool execution failed";
}
function serializeResultBody(ticketId, result) {
    try {
        const serializedBody = JSON.stringify({ ticketId, result });
        if (typeof serializedBody === "string") {
            const byteLength = utf8ByteLength(serializedBody);
            if (byteLength <= MAX_RESULT_BODY_BYTES) {
                return { result, serializedBody, byteLength };
            }
        }
    }
    catch {
        // Replace a post-execution serialization failure with a bounded result.
    }
    const safeResult = {
        ok: false,
        category: "Result",
        error: "Companion result is too large",
    };
    const serializedBody = JSON.stringify({ ticketId, result: safeResult });
    return {
        result: safeResult,
        serializedBody,
        byteLength: utf8ByteLength(serializedBody),
    };
}
function canonicalJson(value) {
    return JSON.stringify(canonicalize(value)) ?? "undefined";
}
function canonicalize(value) {
    if (Array.isArray(value))
        return value.map((item) => canonicalize(item));
    if (!isRecord(value))
        return value;
    const normalized = {};
    for (const key of Object.keys(value).sort()) {
        normalized[key] = canonicalize(value[key]);
    }
    return normalized;
}
async function awaitWithSignal(operation, signal) {
    if (signal.aborted) {
        void Promise.resolve(operation).catch(() => {
            // Observe a non-compliant operation after cancellation.
        });
        throw new BridgeCancelledError();
    }
    return new Promise((resolve, reject) => {
        let settled = false;
        const cleanup = () => signal.removeEventListener("abort", onAbort);
        const finish = (callback) => {
            if (settled)
                return;
            settled = true;
            cleanup();
            callback();
        };
        const onAbort = () => finish(() => reject(new BridgeCancelledError()));
        signal.addEventListener("abort", onAbort, { once: true });
        Promise.resolve(operation).then((value) => finish(() => resolve(value)), (error) => finish(() => reject(error)));
    });
}
async function readJsonObject(response) {
    const declaredLength = response.headers?.get?.("content-length");
    if (declaredLength !== null && declaredLength !== undefined) {
        const normalizedLength = declaredLength.trim();
        if (!/^\d+$/.test(normalizedLength))
            throw new Error("Companion response is too large");
        const declaredBytes = Number(normalizedLength);
        if (!Number.isSafeInteger(declaredBytes) || declaredBytes > MAX_RESPONSE_BYTES) {
            throw new Error("Companion response is too large");
        }
    }
    const text = await readBoundedResponseText(response);
    if (!text)
        return {};
    const value = JSON.parse(text);
    if (!isRecord(value))
        throw new Error("Companion response must be a JSON object");
    return value;
}
async function readBoundedResponseText(response) {
    if (!response.body) {
        const text = await response.text();
        if (utf8ByteLength(text) > MAX_RESPONSE_BYTES) {
            throw new Error("Companion response is too large");
        }
        return text;
    }
    const reader = response.body.getReader();
    const chunks = [];
    let totalBytes = 0;
    try {
        while (true) {
            const { done, value } = await reader.read();
            if (done)
                break;
            if (!value)
                continue;
            totalBytes += value.byteLength;
            if (totalBytes > MAX_RESPONSE_BYTES) {
                await reader.cancel();
                throw new Error("Companion response is too large");
            }
            chunks.push(value);
        }
    }
    finally {
        reader.releaseLock();
    }
    const bytes = new Uint8Array(totalBytes);
    let offset = 0;
    for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.byteLength;
    }
    return new TextDecoder().decode(bytes);
}
function utf8ByteLength(value) {
    return new TextEncoder().encode(value).byteLength;
}
function isTerminalResultRejection(error) {
    if (!(error instanceof BridgeHttpError) || !error.responseBody)
        return false;
    if (error.responseBody.ok !== false)
        return false;
    const category = error.responseBody.category;
    return (error.status === 400 && category === "Validation") ||
        (error.status === 409 && category === "Conflict") ||
        (error.status === 499 && category === "Cancellation");
}
function isRetryableResultError(error) {
    if (error instanceof BridgeCancelledError || error instanceof BridgeProtocolError)
        return false;
    if (isTerminalResultRejection(error))
        return false;
    return true;
}
function isRetryableRegistrationError(error) {
    if (error instanceof ApprovalRequiredError || error instanceof BridgeCancelledError || error instanceof BridgeValidationError)
        return false;
    if (error instanceof BridgeHttpError) {
        return error.status === 401 ||
            error.status === 404 ||
            error.status === 408 ||
            error.status === 425 ||
            error.status === 429 ||
            error.status >= 500;
    }
    return error instanceof BridgeTimeoutError || error instanceof TypeError || error instanceof Error;
}
function isRetryableSessionError(error) {
    if (error instanceof BridgeCancelledError || error instanceof BridgeCapacityError || error instanceof BridgeValidationError)
        return false;
    if (error instanceof BridgeHttpError) {
        return error.status === 401 ||
            error.status === 404 ||
            error.status === 408 ||
            error.status === 425 ||
            error.status === 429 ||
            error.status >= 500;
    }
    return error instanceof BridgeTimeoutError || error instanceof TypeError || error instanceof Error;
}
function backgroundErrorMessage(error) {
    if (error instanceof BridgeTimeoutError)
        return "Companion request timed out";
    if (error instanceof BridgeHttpError)
        return `Companion request failed (${error.status})`;
    return "PowerPoint bridge unavailable";
}
function clampRandom(value) {
    if (!Number.isFinite(value) || value <= 0)
        return 0;
    if (value >= 1)
        return Number.MIN_VALUE + (1 - Number.EPSILON);
    return value;
}
function isRecord(value) {
    return typeof value === "object" && value !== null && !Array.isArray(value);
}
function officeBridge_isUuid(value) {
    return typeof value === "string" &&
        /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

;// ./src/officeBridgePowerPoint.ts



const DOCUMENT_ID_KEY = "hermes.officeBridge.documentId.v1";
async function getOrCreatePowerPointDocumentId(randomUuid = () => crypto.randomUUID()) {
    if (typeof Office === "undefined" || !Office.context?.document?.settings) {
        throw new Error("Office document settings are unavailable");
    }
    const settings = Office.context.document.settings;
    const existing = settings.get(DOCUMENT_ID_KEY);
    if (typeof existing === "string" && officeBridgePowerPoint_isUuid(existing)) {
        return existing;
    }
    const documentId = randomUuid();
    if (!officeBridgePowerPoint_isUuid(documentId)) {
        throw new Error("Generated documentId must be a UUID");
    }
    settings.set(DOCUMENT_ID_KEY, documentId);
    await new Promise((resolve, reject) => {
        settings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve();
            }
            else {
                reject(new Error("Unable to persist the PowerPoint bridge documentId"));
            }
        });
    });
    return documentId;
}
async function executeValidatedPowerPointBridgeTool(name, argumentsValue) {
    const validation = validateToolCall(name, JSON.stringify(argumentsValue));
    if (!("valid" in validation) || validation.valid !== true) {
        return JSON.stringify(validation);
    }
    return executePowerPointToolParsed(name, validation.parsedArgs);
}
function createPowerPointCompanionBridge(onState, desk) {
    return new OfficeCompanionBridge({
        hostPrefix: "powerpoint",
        paneId: crypto.randomUUID(),
        documentId: () => getOrCreatePowerPointDocumentId(),
        executeTool: executeValidatedPowerPointBridgeTool,
        onState,
        baseUrl: desk?.baseUrl,
        grant: desk?.grant,
        onApprovalRequired: desk?.onApprovalRequired,
    });
}
function officeBridgePowerPoint_isUuid(value) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

;// ./src/taskpane/taskpane.ts











const agent = new HermesAgent();
let abortController = null;
let currentAssistantMessage = null;
let assistantBuffer = "";
let settings = loadSettings();
const checkpoint = new CheckpointController();
let bridgeBase = null;
let connecting = false;
let bridge = null;
const connector = new DeskConnector({ client: { host: "PowerPoint", platform: officePlatform() } });
const bridgeLifecycle = new OfficeBridgeLifecycleController((onState) => {
    bridge = createPowerPointCompanionBridge(onState, {
        baseUrl: bridgeBase ?? "",
        grant: () => connector.grant,
        onApprovalRequired: () => void approvalLost(),
    });
    return bridge;
}, updateBridgeState);
function officePlatform() {
    try {
        return String(Office?.context?.platform ?? "");
    }
    catch {
        return "";
    }
}
Office.onReady(() => {
    window.OmniaTheme?.observeOffice();
    bindControls();
    updatePrivacyBadge();
    setStatus("Ready");
    updateBridgeState({ connected: false, sessionId: null });
    addAssistantMessage("Hi, I'm Hermes. Ask me to inspect, draft, or edit the active presentation.");
    bridgeLifecycle.register(window, Office.addin && typeof Office.addin.onVisibilityModeChanged === "function"
        ? {
            hiddenValue: Office.VisibilityMode.hidden,
            register: (handler) => Office.addin.onVisibilityModeChanged((message) => {
                if (bridgeBase)
                    handler(message);
            }),
        }
        : undefined);
    void connectToHermes(false);
});
/**
 * Find the Sidecar bridge, then start a session with the remembered approval. With no approval yet,
 * show "Connect to Hermes"; when `ask` is true (the user clicked it), file the request and wait for the
 * user's Allow/Deny in Hermes.
 */
async function connectToHermes(ask) {
    if (connecting)
        return;
    connecting = true;
    try {
        setConnectPanel("searching");
        const base = await connector.discover();
        if (!base) {
            setConnectPanel(connector.grant ? "error" : "no-bridge");
            return;
        }
        bridgeBase = base;
        settings = { ...settings, baseUrl: `${base}/v1` };
        updatePrivacyBadge();
        if (!connector.grant) {
            if (!ask) {
                setConnectPanel("needs-approval");
                return;
            }
            const outcome = await connector.requestApproval(base, (code, command) => setConnectPanel("waiting", code, command));
            if (outcome !== "approved") {
                setConnectPanel(outcome === "denied" ? "denied" : "expired");
                return;
            }
        }
        await bridgeLifecycle.stop();
        await bridgeLifecycle.start();
        setConnectPanel("connected");
    }
    catch (error) {
        if (error instanceof ApprovalRequiredError) {
            setConnectPanel("needs-approval");
        }
        else {
            setConnectPanel("error");
        }
    }
    finally {
        connecting = false;
    }
}
async function approvalLost() {
    connector.clearGrant();
    await bridgeLifecycle.stop().catch(() => undefined);
    setConnectPanel("needs-approval");
}
function setConnectPanel(state, detail = "", command = DESK_COMMAND) {
    const panel = byId("desk-connect");
    const text = byId("desk-connect-text");
    const button = byId("desk-connect-button");
    const code = byId("desk-connect-code");
    panel.hidden = state === "connected";
    code.hidden = state !== "waiting";
    button.disabled = state === "searching" || state === "waiting";
    button.textContent = state === "no-bridge" || state === "error" ? "Try again" : "Connect to Hermes";
    const messages = {
        "searching": "Looking for Hermes on this computer…",
        "no-bridge": "Hermes is not reachable. Start the Hermes gateway with the API server enabled and the Sidecar for PowerPoint plugin installed, trust its local CA, then try again.",
        "needs-approval": "This sidebar needs your OK in Hermes before it can chat.",
        "waiting": `In any Hermes chat, run ${command} and choose Allow. Hermes shows the same code:`,
        "denied": "Hermes denied this sidebar. Click Connect to Hermes to ask again.",
        "expired": "The request expired. Click Connect to Hermes to ask again.",
        "connected": "",
        "error": "Lost the connection to Hermes on this computer.",
    };
    text.textContent = messages[state];
    code.textContent = state === "waiting" ? detail : "";
    setStatus(state === "connected" ? "Connected to Hermes" : state === "waiting" ? "Waiting for Hermes approval" : "Not connected");
}
function bindControls() {
    const sendButton = byId("send-button");
    const stopButton = byId("stop-button");
    const input = byId("prompt-input");
    const settingsButton = byId("settings-button");
    const settingsDialog = byId("settings-dialog");
    const saveSettingsButton = byId("save-settings-button");
    const copyBridgeButton = byId("copy-bridge-session");
    sendButton.addEventListener("click", () => void sendPrompt());
    stopButton.addEventListener("click", stopRunningAgent);
    settingsButton.addEventListener("click", () => openSettings());
    saveSettingsButton.addEventListener("click", () => persistSettingsFromDialog());
    copyBridgeButton.addEventListener("click", () => void copyBridgeSessionId());
    byId("desk-connect-button").addEventListener("click", () => void connectToHermes(true));
    input.addEventListener("keydown", (event) => {
        if (event.key === "Enter" && !event.shiftKey) {
            event.preventDefault();
            void sendPrompt();
        }
    });
    settingsDialog.addEventListener("cancel", () => fillSettingsDialog());
    byId("checkpoint-continue").addEventListener("click", () => { checkpoint.choose("continue"); byId("checkpoint-bar").classList.add("hidden"); });
    byId("checkpoint-stop").addEventListener("click", () => { checkpoint.choose("stop"); byId("checkpoint-bar").classList.add("hidden"); });
}
// ---------------------------------------------------------------------------
// Companion bridge lifecycle
// ---------------------------------------------------------------------------
function updateBridgeState(state) {
    const status = byId("bridge-status");
    const session = byId("bridge-session-id");
    const copy = byId("copy-bridge-session");
    if (state.connected && state.sessionId) {
        setConnectPanel("connected");
        status.textContent = "Office bridge connected";
        status.className = "bridge-status bridge-connected";
        session.textContent = state.sessionId;
        session.title = state.sessionId;
        copy.disabled = false;
    }
    else {
        status.textContent = state.lastError
            ? "Office bridge unavailable"
            : "Office bridge connecting";
        if (state.lastError)
            setConnectPanel(connector.grant ? "error" : "needs-approval");
        status.className = "bridge-status bridge-disconnected";
        session.textContent = "";
        session.removeAttribute("title");
        copy.disabled = true;
    }
}
async function copyBridgeSessionId() {
    const sessionId = bridgeLifecycle.sessionId;
    if (!sessionId)
        return;
    try {
        await navigator.clipboard.writeText(sessionId);
        setStatus("Bridge session ID copied");
    }
    catch {
        setStatus("Unable to copy bridge session ID");
    }
}
async function sendPrompt() {
    const input = byId("prompt-input");
    const prompt = input.value.trim();
    if (!prompt || abortController) {
        return;
    }
    const chatHeaders = bridge?.chatHeaders() ?? null;
    if (!bridgeBase || !chatHeaders) {
        setConnectPanel(!connector.grant ? "needs-approval" : "error");
        return;
    }
    const bridgeSessionId = bridgeLifecycle.sessionId;
    input.value = "";
    addUserMessage(prompt);
    assistantBuffer = "";
    currentAssistantMessage = addAssistantMessage("");
    setRunning(true);
    abortController = new AbortController();
    const trace = startTrace(settings.maxIterations);
    try {
        const runResult = await agent.run(prompt, { ...settings, requestHeaders: chatHeaders }, {
            onAssistantStart: () => {
                assistantBuffer = "";
                renderAssistantBuffer();
            },
            onAssistantDelta: (delta) => {
                assistantBuffer += delta;
                renderAssistantBuffer();
            },
            onAssistantFinal: (content) => {
                assistantBuffer = content;
                renderAssistantBuffer();
            },
            onActivity: (item) => renderActivity(item),
            onStatus: setStatus,
            onModelCall: (event) => recordModelCall(trace, event),
            onToolResult: (event) => recordToolResult(trace, event.name, event.durationMs, event.ok, event.category),
            onCheckpointResult: (decision) => recordCheckpoint(trace, decision),
            onCheckpoint: async (deadlineSignal) => {
                const bar = byId("checkpoint-bar");
                bar.classList.remove("hidden");
                try {
                    return await checkpoint.begin(deadlineSignal);
                }
                finally {
                    bar.classList.add("hidden");
                }
            }
        }, abortController.signal, { bridgeSessionId });
        trace.outcome = runResult.stopReason || "completed";
    }
    catch (error) {
        trace.outcome = error instanceof DOMException ? error.name : "error";
        if (error instanceof DOMException && error.name === "AbortError") {
            assistantBuffer = `${assistantBuffer.trim()}\n\nStopped.`;
            renderAssistantBuffer();
            setStatus("Stopped");
        }
        else {
            const message = formatAgentError(error);
            assistantBuffer = `I couldn't complete that request.\n\n${message}`;
            renderAssistantBuffer();
            setStatus("Error");
        }
    }
    finally {
        checkpoint.cancel();
        byId("checkpoint-bar").classList.add("hidden");
        finalizeTrace(trace);
        abortController = null;
        currentAssistantMessage = null;
        setRunning(false);
    }
}
function formatAgentError(error) {
    const toolError = error && typeof error === "object" && "toolError" in error
        ? error.toolError
        : undefined;
    return asToolError(toolError ?? error).error;
}
function stopRunningAgent() {
    abortController?.abort();
}
function openSettings() {
    fillSettingsDialog();
    byId("settings-dialog").showModal();
}
function fillSettingsDialog() {
    byId("bridge-url").textContent = bridgeBase ?? "not found";
    byId("model-input").value = settings.model;
    byId("max-iterations-input").value = String(settings.maxIterations);
}
function persistSettingsFromDialog() {
    settings = saveSettings({
        baseUrl: settings.baseUrl,
        model: byId("model-input").value,
        maxIterations: Number(byId("max-iterations-input").value)
    });
    updatePrivacyBadge();
    setStatus("Settings saved");
}
function addUserMessage(content) {
    return addMessage("user", content);
}
function addAssistantMessage(content) {
    return addMessage("assistant", content);
}
function addMessage(role, content) {
    const messages = byId("messages");
    const root = document.createElement("article");
    root.className = `message ${role}`;
    const bubble = document.createElement("div");
    bubble.className = "bubble";
    root.appendChild(bubble);
    messages.appendChild(root);
    const uiMessage = { root, bubble };
    if (role === "assistant") {
        renderMarkdown(bubble, content);
    }
    else {
        bubble.textContent = content;
    }
    scrollMessagesToBottom();
    return uiMessage;
}
function renderAssistantBuffer() {
    if (!currentAssistantMessage) {
        return;
    }
    renderMarkdown(currentAssistantMessage.bubble, assistantBuffer || " ");
    scrollMessagesToBottom();
}
function renderMarkdown(target, markdown) {
    const html = marked.parse(markdown, { async: false, breaks: true });
    target.innerHTML = purify.sanitize(html);
}
function renderActivity(item) {
    if (!currentAssistantMessage) {
        return;
    }
    if (!currentAssistantMessage.activity || !currentAssistantMessage.activityList) {
        const activity = document.createElement("details");
        activity.className = "activity";
        activity.open = true;
        const summary = document.createElement("summary");
        summary.textContent = "Activity";
        activity.appendChild(summary);
        const activityList = document.createElement("div");
        activityList.className = "activity-list";
        activity.appendChild(activityList);
        currentAssistantMessage.root.appendChild(activity);
        currentAssistantMessage.activity = activity;
        currentAssistantMessage.activityList = activityList;
    }
    const activityList = currentAssistantMessage.activityList;
    let row = activityList.querySelector(`[data-activity-id="${item.id}"]`);
    if (!row) {
        row = document.createElement("div");
        row.className = "activity-item";
        row.dataset.activityId = item.id;
        row.innerHTML = "<span></span><span></span>";
        activityList.appendChild(row);
    }
    row.className = `activity-item ${item.status === "error" ? "error" : ""}`;
    row.dataset.status = item.status;
    row.children[0].textContent = item.status === "running" ? "\u2026" : item.status === "done" ? "\u2713" : "!";
    row.children[1].textContent = item.label;
    scrollMessagesToBottom();
}
function setRunning(isRunning) {
    byId("send-button").disabled = isRunning;
    byId("stop-button").disabled = !isRunning;
    byId("prompt-input").disabled = isRunning;
}
function setStatus(status) {
    byId("connection-status").textContent = status;
}
function scrollMessagesToBottom() {
    const messages = byId("messages");
    messages.scrollTop = messages.scrollHeight;
}
function byId(id) {
    const element = document.getElementById(id);
    if (!element) {
        throw new Error(`Missing element #${id}`);
    }
    return element;
}
function updatePrivacyBadge() {
    const badge = document.getElementById("privacy-badge");
    if (!badge)
        return;
    const privacy = classifyPrivacy(bridgeBase ?? "http://127.0.0.1");
    badge.textContent = PRIVACY_LABELS[privacy].label;
    badge.title = PRIVACY_LABELS[privacy].tooltip;
    badge.dataset.privacy = privacy;
}

/******/ })()
;