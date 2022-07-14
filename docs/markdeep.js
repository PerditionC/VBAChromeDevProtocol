// (Note: invisible BOM on this line!)
/** 

  Markdeep.js
  Version 1.13

  Copyright 2015-2021, Morgan McGuire, https://casual-effects.com
  All rights reserved.

  -------------------------------------------------------------

  See https://casual-effects.com/markdeep for documentation on how to
  use this script make your plain text documents render beautifully
  in web browsers.

  Markdeep was created by Morgan McGuire. It extends the work of:

   - John Gruber's original Markdown
   - Ben Hollis' Maruku Markdown dialect
   - Michel Fortin's Markdown Extras dialect
   - Ivan Sagalaev's highlight.js
   - Contributors to the above open source projects

  -------------------------------------------------------------
 
  You may use, extend, and redistribute this code under the terms of
  the BSD license at https://opensource.org/licenses/BSD-2-Clause.

  Contains highlight.js (https://github.com/isagalaev/highlight.js) by Ivan
  Sagalaev, which is used for code highlighting. (BSD 3-clause license)

  There is an invisible Byte-Order-Marker at the start of this file to
  ensure that it is processed as UTF-8. Do not remove this character or it
  will break the regular expressions in highlight.js.
*/
/**See https://casual-effects.com/markdeep for @license and documentation.
markdeep.min.js 1.13 (C) 2020 Morgan McGuire 
highlight.min.js 10.5.0 (C) 2020 Ivan Sagalaev https://highlightjs.org */
(function() {
'use strict';

var MARKDEEP_FOOTER = '<div class="markdeepFooter"><i>formatted by <a href="https://casual-effects.com/markdeep" style="color:#999">Markdeep&nbsp;1.13&nbsp;&nbsp;</a></i><div style="display:inline-block;font-size:13px;font-family:\'Times New Roman\',serif;vertical-align:middle;transform:translate(-3px,-1px)rotate(135deg);">&#x2712;</div></div>';

{
// For minification. This is admittedly scary.
var _ = String.prototype;
_.rp = _.replace;
_.ss = _.substring;
if (!_.endsWith) {
    // For IE11
    _.endsWith = function(S, L) {
        if (L === undefined || L > this.length) {
            L = this.length;
        }
        return this.ss(L - S.length, L) === S;
    };
}

// Regular expression version of String.indexOf
_.regexIndexOf = function(regex, startpos) {
    var i = this.ss(startpos || 0).search(regex);
    return (i >= 0) ? (i + (startpos || 0)) : i;
}
}

/** Enable for debugging to view character bounds in diagrams */
var DEBUG_SHOW_GRID = false;

/** Overlay the non-empty characters of the original source in diagrams */
var DEBUG_SHOW_SOURCE = DEBUG_SHOW_GRID;

/** Use to suppress passing through text in diagrams */
var DEBUG_HIDE_PASSTHROUGH = DEBUG_SHOW_SOURCE;

/** In pixels of lines in diagrams */
var STROKE_WIDTH = 2;

/** A box of these denotes a diagram */
var DIAGRAM_MARKER = '*';

// http://stackoverflow.com/questions/1877475/repeat-character-n-times
// ECMAScript 6 has a String.repeat method, but that's not available everywhere
var DIAGRAM_START = Array(5 + 1).join(DIAGRAM_MARKER);

/** attribs are optional */
function entag(tag, content, attribs) {
    return '<' + tag + (attribs ? ' ' + attribs : '') + '>' + content + '</' + tag + '>';
}


function measureFontSize(fontStack) {
    try {
        var canvas = document.createElement('canvas');
        var ctx = canvas.getContext('2d');
        ctx.font = '10pt ' + fontStack;
        return ctx.measureText("M").width;
    } catch (e) {
        // Needed for Firefox include...canvas doesn't work for some reason
        return 10;
    }
}

// IE11 polyfill needed by Highlight.js, from https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/assign#Polyfill
if (typeof Object.assign !== 'function') {
    // Must be writable: true, enumerable: false, configurable: true
    Object.defineProperty(Object, "assign", {
        value: function assign(target, varArgs) { // .length of function is 2
            if (target === null || target === undefined) {
                throw new TypeError('Cannot convert undefined or null to object');
            }
            
            var to = Object(target);
            
            for (var index = 1; index < arguments.length; index++) {
                var nextSource = arguments[index];
                
                if (nextSource !== null && nextSource !== undefined) { 
                    for (var nextKey in nextSource) {
                        // Avoid bugs when hasOwnProperty is shadowed
                        if (Object.prototype.hasOwnProperty.call(nextSource, nextKey)) {
                            to[nextKey] = nextSource[nextKey];
                        }
                    }
                }
            }
            return to;
        },
        writable: true,
        configurable: true
    });
}

// Polyfill for IE11 from https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/includes
if (!String.prototype.includes) {
    String.prototype.includes = function(search, start) {
        if (search instanceof RegExp) {
            throw TypeError('first argument must not be a RegExp');
        } 
        if (start === undefined) { start = 0; }
        return this.indexOf(search, start) !== -1;
    };
}
if (!Array.prototype.includes) {
    Array.prototype.includes = function(search) {
        return !!~this.indexOf(search);
    }
}
    
   
 
// Lucida Console on Windows has capital V's that look like lower case, so don't use it
var codeFontStack = "Menlo,Consolas,monospace";
var codeFontSize  = Math.round(6.5 * 105.1316178 / measureFontSize(codeFontStack)) + '%';// + 'px';

var BODY_STYLESHEET = entag('style', 'body{max-width:680px;' +
    'margin:auto;' +
    'padding:20px;' +
    'text-align:justify;' +
    'line-height:140%;' +
    '-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;font-smoothing:antialiased;' +
    'color:#222;' +
    'font-family:Palatino,Georgia,"Times New Roman",serif}');

var isFirefox = navigator.userAgent.indexOf('Firefox') !== -1 && navigator.userAgent.indexOf('Seamonkey') === -1;
    
/** You can embed your own stylesheet AFTER the <script> tags in your
    file to override these defaults. */
var STYLESHEET = entag('style',
                       // Force background images (except on the body) to print correctly on Chrome and Safari
                       // and remove text shadows, which Chrome can't print and will turn into
                       // boxes
    '@media print{*{-webkit-print-color-adjust:exact;text-shadow:none !important}}' +

    'body{' +
    'counter-reset: h1 paragraph line item list-item' +
    '}' +

    // Avoid header/footer in print to PDF. See https://productforums.google.com/forum/#!topic/chrome/LBMUDtGqr-0
    '@page{margin:0;size:auto}' +

    '#mdContextMenu{position:absolute;background:#383838;cursor:default;border:1px solid #999;color:#fff;padding:4px 0px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Oxygen,Ubuntu,"Helvetica Neue",sans-serif;font-size:85%;font-weight:600;border-radius:4px;box-shadow:0px 3px 10px rgba(0,0,0,35%)}' +
    '#mdContextMenu div{padding:0px 20px}' +
    '#mdContextMenu div:hover{background:#1659d1}' +
                       
    '.md code,.md pre{' +
    'font-family:' + codeFontStack + ';' +
    'font-size:' + codeFontSize + ';' +
    'text-align:left;' +
    'line-height:140%' + 
    '}' +

    '.md .mediumToc code,.md longToc code,.md .shortToc code,.md h1 code,.md h2 code,.md h3 code,.md h4 code,.md h5 code,.md h6 code{font-size:unset}' +

    '.md div.title{' +
    'font-size:26px;' +
    'font-weight:800;' +
    'line-height:120%;' +
    'text-align:center' +
    '}' +

    '.md div.afterTitles{height:10px}' +

    '.md div.subtitle{' +
    'text-align:center' +
    '}' +

    '.md iframe.textinsert, .md object.textinsert,.md iframe:not(.markdeep){display:block;margin-top:10px;margin-bottom:10px;width:100%;height:75vh;border:1px solid #000;border-radius:4px;background:#f5f5f4}' +

    '.md .image{display:inline-block}' +

    '.md img{' +
    'max-width:100%;' +
    'page-break-inside:avoid' +
    '}' +

    // Justification tends to handle URLs and code blocks poorly
    // when inside of a bullet, so disable it there
    '.md li{text-align:left;text-indent:0}' +

    // Make code blocks use 4-space tabs.
    // Set up a line number counter. Do NOT use "overflow: scroll" or it will force scrollbars even when unused on Windows.
    // Don't use text-overflow:ellipsis; which on mac just makes the line short even when scrolled
    '.md pre.listing {width:100%;tab-size:4;-moz-tab-size:4;-o-tab-size:4;counter-reset:line;overflow-x:auto;resize:horizontal}' +

    '.md pre.listing .linenumbers span.line:before{width:30px;margin-left:-28px;font-size:80%;text-align:right;counter-increment:line;' +
    'content:counter(line);display:inline-block;padding-right:13px;margin-right:8px;color:#ccc}' +

     // Force captions on line listings down close and then center them
    '.md div.tilde{' +
    'margin:20px 0 -10px;' +
    'text-align:center' + 
    '}' +

    '.md .imagecaption,.md .tablecaption,.md .listingcaption{' +
    'display:inline-block;' +
    'margin:7px 5px 12px;' +
    'text-align:justify;' +
    'font-style:italic' +
    '}' +
                       
    '.md img.pixel{image-rendering:-moz-crisp-edges;image-rendering:pixelated}' +
                       
    '.md blockquote.fancyquote{' + 
    'margin:25px 0 25px;' +
    'text-align:left;' +
    'line-height:160%' +
    '}' +

    '.md blockquote.fancyquote::before{' +
    'content:"\u201C";' +
    'color:#DDD;' +
    'font-family:Times New Roman;' +
    'font-size:45px;' +
    'line-height:0;' +
    'margin-right:6px;' +
    'vertical-align:-0.3em' +
    '}' +

    '.md span.fancyquote{' +
    'font-size:118%;' +
    'color:#777;' +
    'font-style:italic' +
    '}' +

    '.md span.fancyquote::after{' +
    'content:"\u201D";' +
    'font-style:normal;' +
    'color:#DDD;' +
    'font-family:Times New Roman;' +
    'font-size:45px;' +
    'line-height:0;' +
    'margin-left:6px;' +
    'vertical-align:-0.3em' +
    '}' +

    '.md blockquote.fancyquote .author{' +
    'width:100%;' +
    'margin-top:10px;' + 
    'display:inline-block;' +
    'text-align:right' +
    '}' +

    '.md small{font-size:60%}' +
    '.md big{font-size:150%}' +

    '.md div.title,contents,.md .tocHeader,.md h1,.md h2,.md h3,.md h4,.md h5,.md h6,.md .shortTOC,.md .mediumTOC,.nonumberh1,.nonumberh2,.nonumberh3,.nonumberh4,.nonumberh5,.nonumberh6{' +
    'font-family:Verdana,Helvetica,Arial,sans-serif;' +
    'margin:13.4px 0 13.4px;' +
    'padding:15px 0 3px;' +
    'border-top:none;' +
    'clear:both' +
    '}' +
                       
    '.md .tocTop {display:none}' +

    '.md h1,.md h2,.md h3,.md h4,.md h5,.md h6,.md .nonumberh1,.md .nonumberh2,.md .nonumberh3,.md .nonumberh4,.md .nonumberh5,.md .nonumberh6{' +
     'page-break-after:avoid;break-after:avoid' +
    '}'+

    '.md svg.diagram{' +
    'display:block;' +
    'font-family:' + codeFontStack + ';' +
    'font-size:' + codeFontSize + ';' +
    'text-align:center;' +
    'stroke-linecap:round;' +
    'stroke-width:' + STROKE_WIDTH + 'px;'+
    'page-break-inside:avoid;' +
    'stroke:#000;' + 
    'fill:#000' +
    '}' +

    '.md svg.diagram .opendot{' +
    'fill:#fff' +
    '}' +

    '.md svg.diagram .shadeddot{' +
    'fill:#CCC' +
    '}' +

    '.md svg.diagram .dotteddot{' +
    'stroke:#000;stroke-dasharray:4;fill:none' +
    '}' +

    '.md svg.diagram text{' +
    'stroke:none' +
    '}' +

    // printing scale and margins
    '@media print{@page{margin:1in 5mm;transform: scale(150%)}}' +
                       
    // pagebreak hr
    '@media print{.md .pagebreak{page-break-after:always;visibility:hidden}}' +

    // Not restricted to a:link because we want things like svn URLs to have this font, which
    // makes "//" look better.
    '.md a{font-family:Georgia,Palatino,\'Times New Roman\'}' +

    '.md h1,.md .tocHeader,.md .nonumberh1{' +
    'border-bottom:3px solid;' +
    'font-size:20px;' +
    'font-weight:bold;' +
    '}' +

    '.md h1,.md .nonumberh1{' +
    'counter-reset:h2 h3 h4 h5 h6' +
    '}' +

    '.md h2,.md .nonumberh2{' +
    'counter-reset:h3 h4 h5 h6;' +
    'border-bottom:2px solid #999;' +
    'color:#555;' +
    'font-weight:bold;'+
    'font-size:18px;' +
    '}' +

    '.md h3,.md h4,.md h5,.md h6,.md .nonumberh3,.md .nonumberh4,.md .nonumberh5,.md .nonumberh6{' +
    'font-family:Verdana,Helvetica,Arial,sans-serif;' +
    'color:#555;' +
    'font-size:16px;' +
    '}' +

    '.md h3{counter-reset:h4 h5 h6}' +
    '.md h4{counter-reset:h5 h6}' +
    '.md h5{counter-reset:h6}' +

    '.md div.table{' +
    'margin:16px 0 16px 0' +
    '}' +
                       
    '.md table{' +
    'border-collapse:collapse;' +
    'line-height:140%;' +
    'page-break-inside:avoid' +
    '}' +

    '.md table.table{' +
    'margin:auto' +
    '}' +

    '.md table.calendar{' +
    'width:100%;' +
    'margin:auto;' +
    'font-size:11px;' +
    'font-family:Verdana,Helvetica,Arial,sans-serif' +
    '}' +

    '.md table.calendar th{' +
    'font-size:16px' +
    '}' +

    '.md .today{' +
    'background:#ECF8FA' +
    '}' +

    '.md .calendar .parenthesized{' +
    'color:#999;' + 
    'font-style:italic' +
    '}' +

    '.md table.table th{' +
    'color:#FFF;' +
    'background-color:#AAA;' +
    'border:1px solid #888;' +
     // top right bottom left
    'padding:8px 15px 8px 15px' +
    '}' +

    '.md table.table td{' +
     // top right bottom left
    'padding:5px 15px 5px 15px;' +
    'border:1px solid #888' +
    '}' +

    '.md table.table tr:nth-child(even){'+
    'background:#EEE' +
    '}' +

    '.md pre.tilde{' +
    'border-top: 1px solid #CCC;' + 
    'border-bottom: 1px solid #CCC;' + 
    'padding: 5px 0 5px 20px;' +
    'margin:0 0 0 0;' +
    'background:#FCFCFC;' +
    'page-break-inside:avoid' +
    '}' +

    '.md a.target{width:0px;height:0px;visibility:hidden;font-size:0px;display:inline-block}' +
    '.md a:link, .md a:visited{color:#38A;text-decoration:none}' +
    '.md a:link:hover{text-decoration:underline}' +

    '.md dt{' +
    'font-weight:700' +
    '}' +

    // Remove excess space above definitions due to paragraph breaks, and add some at the bottom
    '.md dl>dd{margin-top:-8px; margin-bottom:8px}' +
                       
     // Extra space around terse definition lists
    '.md dl>table{' +
    'margin:35px 0 30px' + 
    '}' +

    '.md code{' +
    'page-break-inside:avoid;' +
    '} @media print{.md .listing code{white-space:pre-wrap}}' +

    '.md .endnote{' +
    'font-size:13px;' +
    'line-height:15px;' +
    'padding-left:10px;' +
    'text-indent:-10px' +
    '}' +

    '.md .bib{' +
    'padding-left:80px;' +
    'text-indent:-80px;' +
    'text-align:left' +
    '}' +

    '.markdeepFooter{font-size:9px;text-align:right;padding-top:80px;color:#999}' +

    '.md .mediumTOC{float:right;font-size:12px;line-height:15px;border-left:1px solid #CCC;padding-left:15px;margin:15px 0px 15px 25px}' +

    '.md .mediumTOC .level1{font-weight:600}' +

    '.md .longTOC .level1{font-weight:600;display:block;padding-top:12px;margin:0 0 -20px}' +
     
    '.md .shortTOC{text-align:center;font-weight:bold;margin-top:15px;font-size:14px}' +

    '.md .admonition{' +
         'position:relative;' +
         'margin:1em 0;' +
         'padding:.4rem 1rem;' +
         'border-radius:.2rem;' +
         'border-left:2.5rem solid rgba(68,138,255,.4);' +
         'background-color:rgba(68,138,255,.15);' +
     '}' +

     '.md .admonition-title{' +
         'font-weight:bold;' +
         'border-bottom:solid 1px rgba(68,138,255,.4);' +
         'padding-bottom:4px;' +
         'margin-bottom:4px;' +
         'margin-left: -1rem;' +
         'padding-left:1rem;' +
         'margin-right:-1rem;' +
         'border-color:rgba(68,138,255,.4)' +
     '}' +

    '.md .admonition.tip{' +
       'border-left:2.5rem solid rgba(50,255,90,.4);' +
       'background-color:rgba(50,255,90,.15)' +
    '}' +
                       
    '.md .admonition.tip::before{' +
       'content:"\\24d8";' +
       'font-weight:bold;' +
       'font-size:' + (isFirefox ? '200%;' : '150%;') +
       'position:relative;' +
       'top:3px;' +
       'color:rgba(26,128,46,.8);' +
       'left:-2.95rem;' +
       'display:block;' +
       'width:0;' +
       'height:0' +
     '}' +

     '.md .admonition.tip>.admonition-title{' +
       'border-color:rgba(50,255,90,.4)' +
     '}' +

     '.md .admonition.warn,.md .admonition.warning{' +
       'border-left:2.5rem solid rgba(255,145,0,.4);' +
       'background-color:rgba(255,145,0,.15)' +
     '}' +

     '.md .admonition.warn::before,.md .admonition.warning::before{' +
       'content:"\\26A0";' +
       'font-weight:bold;' +
       (isFirefox ? '' : 'font-size:150%;') +
       'position:relative;' +
       'top:2px;' +
       'color:rgba(128,73,0,.8);' +
       'left:-2.95rem;' +
       'display:block;' +
       'width:0;' +
       'height:0' +
     '}' +

     '.md .admonition.warn>.admonition-title,.md .admonition.warning>.admonition-title{' +
      'border-color:rgba(255,145,0,.4)' +
     '}' +

     '.md .admonition.error{' +
      'border-left: 2.5rem solid rgba(255,23,68,.4);'+    
      'background-color:rgba(255,23,68,.15)' +
    '}' +

    '.md .admonition.error>.admonition-title{' +
      'border-color:rgba(255,23,68,.4)'+
    '}' +

    '.md .admonition.error::before{' + 
    'content: "\\2612";' +
    'font-family:"Arial";' +
    'font-size:' + (isFirefox ? '150%;' :'200%;') +
    'position:relative;' +
    'color:rgba(128,12,34,.8);' +
    'top:-2px;' +
    'left:-3rem;' +
    'display:block;' +
    'width:0;' +
    'height:0' +
   '}' +
                       
   '.md .admonition p:last-child{margin-bottom:0}'  +

   '.md li.checked,.md li.unchecked{'+
    'list-style:none;'+
    'overflow:visible;'+
    'text-indent:-1.2em'+
                       '}' +
                       
   '.md li.checked:before,.md li.unchecked:before{' +
   'content:"\\2611";' +
   'display:block;'+
   'float:left;' +
   'width:1em;' +
   'font-size:120%'+
                       '}'+
                       
   '.md li.unchecked:before{'+
   'content:"\\2610"' +
   '}'

);

var MARKDEEP_LINE = '<!-- Markdeep: --><style class="fallback">body{visibility:hidden;white-space:pre;font-family:monospace}</style><script src="markdeep.min.js"></script><script src="https://casual-effects.com/markdeep/latest/markdeep.min.js?"></script><script>window.alreadyProcessedMarkdeep||(document.body.style.visibility="visible")</script>';

// Language options:
var FRENCH = {
    keyword: {
        table:     'tableau',
        figure:    'figure',
        listing:   'liste',
        diagram:   'diagramme',
        contents:  'Table des matières',

        sec:       'sec',
        section:   'section',
        subsection: 'paragraphe',
        chapter:   'chapitre',

        Monday:    'lundi',
        Tuesday:   'mardi',
        Wednesday: 'mercredi',
        Thursday:  'jeudi',
        Friday:    'vendredi',
        Saturday:  'samedi',
        Sunday:    'dimanche',

        January:   'Janvier',
        February:  'Février',
        March:     'Mars',
        April:     'Avril',
        May:       'Mai',
        June:      'Juin', 
        July:      'Juillet',
        August:    'Août', 
        September: 'Septembre', 
        October:   'Octobre', 
        November:  'Novembre',
        December:  'Décembre',

        jan: 'janv.',
        feb: 'févr.',
        mar: 'mars',
        apr: 'avril',
        may: 'mai',
        jun: 'juin',
        jul: 'juil.',
        aug: 'août',
        sep: 'sept.',
        oct: 'oct.',
        nov: 'nov.',
        dec: 'déc.',

        '&ldquo;': '&laquo;&nbsp;',
        '&rtquo;': '&nbsp;&raquo;'
    }
};

// Translated by "Warmist"
var LITHUANIAN = {
    keyword: {
        table:     'lentelė',
        figure:    'paveikslėlis',
        listing:   'sąrašas',
        diagram:   'diagrama',
        contents:  'Turinys',

        sec:       'sk',
        section:   'skyrius',
        subsection: 'poskyris',
        chapter:   'skyrius',

        Monday:    'pirmadienis',
        Tuesday:   'antradienis',
        Wednesday: 'trečiadienis',
        Thursday:  'ketvirtadienis',
        Friday:    'penktadienis',
        Saturday:  'šeštadienis',
        Sunday:    'sekmadienis',

        January:   'Sausis',
        February:  'Vasaris',
        March:     'Kovas',
        April:     'Balandis',
        May:       'Gegužė',
        June:      'Birželis',
        July:      'Liepa',
        August:    'Rugpjūtis',
        September: 'Rugsėjis',
        October:   'Spalis',
        November:  'Lapkritis',
        December:  'Gruodis',

        jan: 'saus',
        feb: 'vas',
        mar: 'kov',
        apr: 'bal',
        may: 'geg',
        jun: 'birž',
        jul: 'liep',
        aug: 'rugpj',
        sep: 'rugs',
        oct: 'spal',
        nov: 'lapkr',
        dec: 'gruod',

        '&ldquo;': '&bdquo;',
        '&rtquo;': '&ldquo;'
    }
};

    
// Translated by Zdravko Velinov
var BULGARIAN = {
    keyword: {
        table:     'таблица',
        figure:    'фигура',
        listing:   'списък',
        diagram:   'диаграма',

        contents:  'cъдържание',

        sec:       'сек',
        section:   'раздел',
        subsection: 'подраздел',
        chapter:   'глава',

        Monday:    'понеделник',
        Tuesday:   'вторник',
        Wednesday: 'сряда',
        Thursday:  'четвъртък',
        Friday:    'петък',
        Saturday:  'събота',
        Sunday:    'неделя',

        January:   'януари',
        February:  'февруари',
        March:     'март',
        April:     'април',
        May:       'май',
        June:      'юни', 
        July:      'юли',
        August:    'август', 
        September: 'септември', 
        October:   'октомври', 
        November:  'ноември',
        December:  'декември',

        jan: 'ян',
        feb: 'февр',
        mar: 'март',
        apr: 'апр',
        may: 'май',
        jun: 'юни',
        jul: 'юли',
        aug: 'авг',
        sep: 'септ',
        oct: 'окт',
        nov: 'ноем',
        dec: 'дек',

        '&ldquo;': '&bdquo;',
        '&rdquo;': '&rdquo;'
    }
};


// Translated by Tiago Antão
var PORTUGUESE = {
    keyword: {
        table:     'tabela',
        figure:    'figura',
        listing:   'lista',
        diagram:   'diagrama',
        contents:  'conteúdo',

        sec:       'sec',
        section:   'secção',
        subsection: 'subsecção',
        chapter:   'capítulo',

        Monday:    'Segunda-feira',
        Tuesday:   'Terça-feira',
        Wednesday: 'Quarta-feira',
        Thursday:  'Quinta-feira',
        Friday:    'Sexta-feira',
        Saturday:  'Sábado',
        Sunday:    'Domingo',

        January:   'Janeiro',
        February:  'Fevereiro',
        March:     'Março',
        April:     'Abril',
        May:       'Maio',
        June:      'Junho', 
        July:      'Julho',
        August:    'Agosto', 
        September: 'Setembro', 
        October:   'Outubro', 
        November:  'Novembro',
        December:  'Dezembro',

        jan: 'jan',
        feb: 'fev',
        mar: 'mar',
        apr: 'abr',
        may: 'mai',
        jun: 'jun',
        jul: 'jul',
        aug: 'ago',
        sep: 'set',
        oct: 'oct',
        nov: 'nov',
        dec: 'dez',

        '&ldquo;': '&laquo;',
        '&rtquo;': '&raquo;'
    }
};


// Translated by Jan Toušek
var CZECH = {
    keyword: {
        table:     'Tabulka',
        figure:    'Obrázek',
        listing:   'Seznam',
        diagram:   'Diagram',

        contents:  'Obsah',

        sec:       'kap.',  // Abbreviation for section
        section:   'kapitola',
        subsection:'podkapitola',
        chapter:   'kapitola',

        Monday:    'pondělí',
        Tuesday:   'úterý',
        Wednesday: 'středa',
        Thursday:  'čtvrtek',
        Friday:    'pátek',
        Saturday:  'sobota',
        Sunday:    'neděle',

        January:   'leden',
        February:  'únor',
        March:     'březen',
        April:     'duben',
        May:       'květen',
        June:      'červen',
        July:      'červenec',
        August:    'srpen',
        September: 'září',
        October:   'říjen',
        November:  'listopad',
        December:  'prosinec',

        jan: 'led',
        feb: 'úno',
        mar: 'bře',
        apr: 'dub',
        may: 'kvě',
        jun: 'čvn',
        jul: 'čvc',
        aug: 'srp',
        sep: 'zář',
        oct: 'říj',
        nov: 'lis',
        dec: 'pro',

        '&ldquo;': '&bdquo;',
        '&rdquo;': '&ldquo;'
    }
};


var ITALIAN = {
    keyword: {
        table:     'tabella',
        figure:    'figura',
        listing:   'lista',
        diagram:   'diagramma',
        contents:  'indice',

        sec:       'sez',
        section:   'sezione',
        subsection: 'paragrafo',
        chapter:   'capitolo',

        Monday:    'lunedì',
        Tuesday:   'martedì',
        Wednesday: 'mercoledì',
        Thursday:  'giovedì',
        Friday:    'venerdì',
        Saturday:  'sabato',
        Sunday:    'domenica',

        January:   'Gennaio',
        February:  'Febbraio',
        March:     'Marzo',
        April:     'Aprile',
        May:       'Maggio',
        June:      'Giugno', 
        July:      'Luglio',
        August:    'Agosto', 
        September: 'Settembre', 
        October:   'Ottobre', 
        November:  'Novembre',
        December:  'Dicembre',

        jan: 'gen',
        feb: 'feb',
        mar: 'mar',
        apr: 'apr',
        may: 'mag',
        jun: 'giu',
        jul: 'lug',
        aug: 'ago',
        sep: 'set',
        oct: 'ott',
        nov: 'nov',
        dec: 'dic',

        '&ldquo;': '&ldquo;',
        '&rtquo;': '&rdquo;'
    }
};

var RUSSIAN = {
    keyword: {
        table:     'таблица',
        figure:    'рисунок',
        listing:   'листинг',
        diagram:   'диаграмма',

        contents:  'Содержание',

        sec:       'сек',
        section:   'раздел',
        subsection: 'подраздел',
        chapter:   'глава',

        Monday:    'понедельник',
        Tuesday:   'вторник',
        Wednesday: 'среда',
        Thursday:  'четверг',
        Friday:    'пятница',
        Saturday:  'суббота',
        Sunday:    'воскресенье',

        January:   'январьr',
        February:  'февраль',
        March:     'март',
        April:     'апрель',
        May:       'май',
        June:      'июнь', 
        July:      'июль',
        August:    'август', 
        September: 'сентябрь', 
        October:   'октябрь', 
        November:  'ноябрь',
        December:  'декабрь',

        jan: 'янв',
        feb: 'февр',
        mar: 'март',
        apr: 'апр',
        may: 'май',
        jun: 'июнь',
        jul: 'июль',
        aug: 'авг',
        sep: 'сент',
        oct: 'окт',
        nov: 'ноябрь',
        dec: 'дек',
        
        '&ldquo;': '«',
        '&rdquo;': '»'
    }
};

// Translated by Dariusz Kuśnierek 
var POLISH = {
    keyword: {
        table:     'tabela',
        figure:    'ilustracja',
        listing:   'wykaz',
        diagram:   'diagram',
        contents:  'Spis treści',

        sec:       'rozdz.',
        section:   'rozdział',
        subsection: 'podrozdział',
        chapter:   'kapituła',

        Monday:    'Poniedziałek',
        Tuesday:   'Wtorek',
        Wednesday: 'Środa',
        Thursday:  'Czwartek',
        Friday:    'Piątek',
        Saturday:  'Sobota',
        Sunday:    'Niedziela',

        January:   'Styczeń',
        February:  'Luty',
        March:     'Marzec',
        April:     'Kwiecień',
        May:       'Maj',
        June:      'Czerwiec', 
        July:      'Lipiec',
        August:    'Sierpień', 
        September: 'Wrzesień', 
        October:   'Październik', 
        November:  'Listopad',
        December:  'Grudzień',

        jan: 'sty',
        feb: 'lut',
        mar: 'mar',
        apr: 'kwi',
        may: 'maj',
        jun: 'cze',
        jul: 'lip',
        aug: 'sie',
        sep: 'wrz',
        oct: 'paź',
        nov: 'lis',
        dec: 'gru',
        
        '&ldquo;': '&bdquo;',
        '&rdquo;': '&rdquo;'
    }
};

// Translated by Sandor Berczi
var HUNGARIAN = {
    keyword: {
        table:     'táblázat',
        figure:    'ábra',
        listing:   'lista',
        diagram:   'diagramm',

        contents:  'Tartalomjegyzék',

        sec:       'fej',  // Abbreviation for section
        section:   'fejezet',
        subsection:'alfejezet',
        chapter:   'fejezet',

        Monday:    'hétfő',
        Tuesday:   'kedd',
        Wednesday: 'szerda',
        Thursday:  'csütörtök',
        Friday:    'péntek',
        Saturday:  'szombat',
        Sunday:    'vasárnap',

        January:   'január',
        February:  'február',
        March:     'március',
        April:     'április',
        May:       'május',
        June:      'június',
        July:      'július',
        August:    'augusztus',
        September: 'szeptember',
        October:   'október',
        November:  'november',
        December:  'december',

        jan: 'jan',
        feb: 'febr',
        mar: 'márc',
        apr: 'ápr',
        may: 'máj',
        jun: 'jún',
        jul: 'júl',
        aug: 'aug',
        sep: 'szept',
        oct: 'okt',
        nov: 'nov',
        dec: 'dec',

        '&ldquo;': '&bdquo;',
        '&rdquo;': '&rdquo;'
    }
};

// Translated by Takashi Masuyama
var JAPANESE = {
    keyword: {
        table:     '表',
        figure:    '図',
        listing:   '一覧',
        diagram:   '図',
        contents:  '目次',

        sec:       '節',
        section:   '節',
        subsection: '項',
        chapter:   '章',

        Monday:    '月',
        Tuesday:   '火',
        Wednesday: '水',
        Thursday:  '木',
        Friday:    '金',
        Saturday:  '土',
        Sunday:    '日',

        January:   '1月',
        February:  '2月',
        March:     '3月',
        April:     '4月',
        May:       '5月',
        June:      '6月',
        July:      '7月',
        August:    '8月',
        September: '9月',
        October:   '10月',
        November:  '11月',
        December:  '12月',

        jan: '1月',
        feb: '2月',
        mar: '3月',
        apr: '4月',
        may: '5月',
        jun: '6月',
        jul: '7月',
        aug: '8月',
        sep: '9月',
        oct: '10月',
        nov: '11月',
        dec: '12月',

        '&ldquo;': '「',
        '&rdquo;': '」'
    }
};    
    
// Translated by Sandor Berczi
var GERMAN = {
    keyword: {
        table:     'Tabelle',
        figure:    'Abbildung',
        listing:   'Auflistung',
        diagram:   'Diagramm',

        contents:  'Inhaltsverzeichnis',

        sec:       'Kap',
        section:   'Kapitel',
        subsection:'Unterabschnitt',
        chapter:   'Kapitel',

        Monday:    'Montag',
        Tuesday:   'Dienstag',
        Wednesday: 'Mittwoch',
        Thursday:  'Donnerstag',
        Friday:    'Freitag',
        Saturday:  'Samstag',
        Sunday:    'Sonntag',

        January:   'Januar',
        February:  'Februar',
        March:     'März',
        April:     'April',
        May:       'Mai',
        June:      'Juni',
        July:      'Juli',
        August:    'August',
        September: 'September',
        October:   'Oktober',
        November:  'November',
        December:  'Dezember',

        jan: 'Jan',
        feb: 'Feb',
        mar: 'Mär',
        apr: 'Apr',
        may: 'Mai',
        jun: 'Jun',
        jul: 'Jul',
        aug: 'Aug',
        sep: 'Sep',
        oct: 'Okt',
        nov: 'Nov',
        dec: 'Dez',
        
        '&ldquo;': '&bdquo;',
        '&rdquo;': '&ldquo;'
    }
};

// Translated by Marcelo Arroyo
var SPANISH = {
    keyword: {
        table:     'Tabla',
        figure:    'Figura',
        listing:   'Listado',
        diagram:   'Diagrama',
        contents:  'Tabla de Contenidos',

        sec:       'sec',
        section:   'Sección',
        subsection: 'Subsección',
        chapter:    'Capítulo',

        Monday:    'Lunes',
        Tuesday:   'Martes',
        Wednesday: 'Miércoles',
        Thursday:  'Jueves',
        Friday:    'Viernes',
        Saturday:  'Sábado',
        Sunday:    'Domingo',

        January:   'Enero',
        February:  'Febrero',
        March:     'Marzo',
        April:     'Abril',
        May:       'Mayo',
        June:      'Junio',
        July:      'Julio',
        August:    'Agosto',
        September: 'Septiembre',
        October:   'Octubre',
        November:  'Noviembre',
        December:  'Diciembre',

        jan: 'ene',
        feb: 'feb',
        mar: 'mar',
        apr: 'abr',
        may: 'may',
        jun: 'jun',
        jul: 'jul',
        aug: 'ago',
        sep: 'sept',
        oct: 'oct',
        nov: 'nov',
        dec: 'dic',

        '&ldquo;': '&laquo;&nbsp;',
        '&rtquo;': '&nbsp;&raquo;'
    }
};

// Translated by Nils Nilsson
var SWEDISH = {
    keyword: {
        table:     'tabell',
        figure:    'figur',
        listing:   'lista',
        diagram:   'diagram',

        contents:  'Innehållsförteckning',
        sec:       'sek',
        section:   'sektion',
        subsection:'sektion',
        chapter:   'kapitel',

        Monday:    'måndag',
        Tuesday:   'tisdag',
        Wednesday: 'onsdag',
        Thursday:  'torsdag',
        Friday:    'fredag',
        Saturday:  'lördag',
        Sunday:    'söndag',

        January:   'januari',
        February:  'februari',
        March:     'mars',
        April:     'april',
        May:       'maj',
        June:      'juni',
        July:      'juli',
        August:    'augusti',
        September: 'september',
        October:   'oktober',
        November:  'november',
        December:  'december',

        jan: 'jan',
        feb: 'feb',
        mar: 'mar',
        apr: 'apr',
        may: 'maj',
        jun: 'jun',
        jul: 'jul',
        aug: 'aug',
        sep: 'sep',
        oct: 'okt',
        nov: 'nov',
        dec: 'dec',
        
        '&ldquo;': '&rdquo;',
        '&rdquo;': '&rdquo;'
    }
};


// Translated by Marc Izquierdo
var CATALAN = {
    keyword: {
        table:     'Taula',
        figure:    'Figura',
        listing:   'Llistat',
        diagram:   'Diagrama',
        contents:  'Taula de Continguts',

        sec:        'sec',
        section:    'Secció',
        subsection: 'Subsecció',
        chapter:    'Capítol',

        Monday:    'Dilluns',
        Tuesday:   'Dimarts',
        Wednesday: 'Dimecres',
        Thursday:  'Dijous',
        Friday:    'Divendres',
        Saturday:  'Dissabte',
        Sunday:    'Dimenge',

        January:   'Gener',
        February:  'Febrer',
        March:     'Març',
        April:     'Abril',
        May:       'Maig',
        June:      'Juny',
        July:      'Juliol',
        August:    'Agost',
        September: 'Septembre',
        October:   'Octubre',
        November:  'Novembre',
        December:  'Desembre',

        jan: 'gen',
        feb: 'feb',
        mar: 'mar',
        apr: 'abr',
        may: 'mai',
        jun: 'jun',
        jul: 'jul',
        aug: 'ago',
        sep: 'sept',
        oct: 'oct',
        nov: 'nov',
        dec: 'des',

        '&ldquo;': '&laquo;&nbsp;',
        '&rtquo;': '&nbsp;&raquo;'
    }
};
 
var DEFAULT_OPTIONS = {
    mode:               'markdeep',
    detectMath:         true,
    lang:               {keyword:{}}, // English
    tocStyle:           'auto',
    hideEmptyWeekends:  true,
    autoLinkImages:     true,
    showLabels:         false,
    sortScheduleLists:  true,
    definitionStyle:    'auto',
    linkAPIDefinitions: false,
    inlineCodeLang:     false,
    scrollThreshold:    90,
    captionAbove:       {diagram: false,
                         image:   false,
                         table:   false,
                         listing: false},
    smartQuotes:        true
};


// See http://www.i18nguy.com/unicode/language-identifiers.html for keys
var LANG_TABLE = {
    en: {keyword:{}},        
    ru: RUSSIAN,
    fr: FRENCH,
    pl: POLISH,
    bg: BULGARIAN,
    de: GERMAN,
    hu: HUNGARIAN,
    sv: SWEDISH,
    pt: PORTUGUESE,
    ja: JAPANESE,
    it: ITALIAN,
    lt: LITHUANIAN,
    cz: CZECH,
    es: SPANISH,
    'es-ES': SPANISH,
    'es-ca': CATALAN
    // Contribute your language here! I only accept translations
    // from native speakers.
};

[].slice.call(document.getElementsByTagName('meta')).forEach(function(elt) {
    var att = elt.getAttribute('lang');
    if (att) {
        var lang = LANG_TABLE[att];
        if (lang) {
            DEFAULT_OPTIONS.lang = lang;
        }
    }
});


var max = Math.max;
var min = Math.min;
var abs = Math.abs;
var sign = Math.sign || function (x) {
    return ( +x === x ) ? ((x === 0) ? x : (x > 0) ? 1 : -1) : NaN;
};


/** Get an option, or return the corresponding value from DEFAULT_OPTIONS */
function option(key, key2) {
    if (window.markdeepOptions && (window.markdeepOptions[key] !== undefined)) {
        var val = window.markdeepOptions[key];
        if (key2) {
            val = val[key2]
            if (val !== undefined) {
                return val;
            } else {
                return DEFAULT_OPTIONS[key][key2];
            }
        } else {
            return window.markdeepOptions[key];
        }
    } else if (DEFAULT_OPTIONS[key] !== undefined) {
        if (key2) {
            return DEFAULT_OPTIONS[key][key2];
        } else {
            return DEFAULT_OPTIONS[key];
        }
    } else {
        console.warn('Illegal option: "' + key + '"');
        return undefined;
    }
}


function maybeShowLabel(url, tag) {
    if (option('showLabels')) {
        var text = ' {\u00A0' + url + '\u00A0}';
        return tag ? entag(tag, text) : text;
    } else {
        return '';
    }
}


// Returns the localized version of word, defaulting to the word itself
function keyword(word) {
    return option('lang').keyword[word] || option('lang').keyword[word.toLowerCase()] || word;
}


/** Converts <>&" to their HTML escape sequences */
function escapeHTMLEntities(str) {
    return String(str).rp(/&/g, '&amp;').rp(/</g, '&lt;').rp(/>/g, '&gt;').rp(/"/g, '&quot;');
}


/** Restores the original source string's '<' and '>' as entered in
    the document, before the browser processed it as HTML. There is no
    way in an HTML document to distinguish an entity that was entered
    as an entity. */
function unescapeHTMLEntities(str) {
    // Process &amp; last so that we don't recursively unescape
    // escaped escape sequences.
    return str.
        rp(/&lt;/g, '<').
        rp(/&gt;/g, '>').
        rp(/&quot;/g, '"').
        rp(/&#39;/g, "'").
        rp(/&ndash;/g, '\u2013').
        rp(/&mdash;/g, '---').
        rp(/&amp;/g, '&');
}


function removeHTMLTags(str) {
    return str.rp(/<.*?>/g, '');
}


/** Turn the argument into a legal URL anchor */
function mangle(text) {
    return encodeURI(text.rp(/\s/g, '').toLowerCase());
}

/** Creates a style sheet containing elements like:

  hn::before { 
    content: counter(h1) "." counter(h2) "." ... counter(hn) " "; 
    counter-increment: hn; 
   } 
*/
function sectionNumberingStylesheet() {
    var s = '';

    for (var i = 1; i <= 6; ++i) {
        s += '.md h' + i + '::before {\ncontent:';
        for (var j = 1; j <= i; ++j) {
            s += 'counter(h' + j + ') "' + ((j < i) ? '.' : ' ') + '"';
        }
        s += ';\ncounter-increment: h' + i + ';margin-right:10px}\n\n';
    }

    return entag('style', s);
}

/**
   \param node  A node from an HTML DOM

   \return A String that is a very good reconstruction of what the
   original source looked like before the browser tried to correct
   it to legal HTML.
 */
function nodeToMarkdeepSource(node, leaveEscapes) {
    var source = node ? node.innerHTML : '';

    // Markdown uses <john@bar.com> email syntax, which HTML parsing
    // will try to close by inserting the matching close tags at the end of the
    // document. Remove anything that looks like that and comes *after*
    // the first fallback style.
    //source = source.rp(/<style class="fallback">[\s\S]*?<\/style>/gi, '');
    
    // Remove artificially inserted close tags from URLs and
    source = source.rp(/<\/https?:.*>|<\/ftp:.*>|<\/[^ "\t\n>]+@[^ "\t\n>]+>/gi, '');
    
    // Now try to fix the URLs themselves, which will be 
    // transformed like this: <http: casual-effects.com="" markdeep="">
    source = source.rp(/<(https?|ftp): (.*?)>/gi, function (match, protocol, list) {

        // Remove any quotes--they wouldn't have been legal in the URL anyway
        var s = '<' + protocol + '://' + list.rp(/=""\s/g, '/');

        if (s.ss(s.length - 3) === '=""') {
            s = s.ss(0, s.length - 3);
        }

        // Remove any lingering quotes (since they
        // wouldn't have been legal in the URL)
        s = s.rp(/"/g, '');

        return s + '>';
    });

    // Remove the "fallback" style tags
    source = source.rp(/<style class=["']fallback["']>.*?<\/style>/gmi, '');

    source = unescapeHTMLEntities(source);

    return source;
}


/** Extracts one diagram from a Markdown string.

    Returns {beforeString, diagramString, alignmentHint, afterString}
    diagramString will be empty if nothing was found. The
    DIAGRAM_MARKER is stripped from the diagramString. 

    alignmentHint may be:
    floatleft  
    floatright
    center
    flushleft

    diagramString does not include the marker characters. 
    If there is a caption, it will appear in the afterString and not be parsed.
*/
function extractDiagram(sourceString) {
    // Returns the number of wide Unicode symbols (outside the BMP) in string s between indices
    // start and end - 1
    function unicodeSyms(s, start, end) {
        var p = start;
        for (var i = start; i < end; ++i, ++p) {
            var c = s.charCodeAt(p);
            p += (c >= 0xD800) && (c <= 0xDBFF);
        }
        return p - end;
    }

    function advance() {
        nextLineBeginning = sourceString.indexOf('\n', lineBeginning) + 1;
        wideCharacters = unicodeSyms(sourceString, lineBeginning + xMin, lineBeginning + xMax);
        textOnLeft  = textOnLeft  || /\S/.test(sourceString.ss(lineBeginning, lineBeginning + xMin));
        noRightBorder = noRightBorder || (sourceString[lineBeginning + xMax + wideCharacters] !== '*');

        // Text on the right ... if the line is not all '*'
        textOnRight = ! noRightBorder && (textOnRight || /[^ *\t\n\r]/.test(sourceString.ss(lineBeginning + xMax + wideCharacters + 1, nextLineBeginning)));
    }

    var noDiagramResult = {beforeString: sourceString, diagramString: '', alignmentHint: '', afterString: ''};

    // Search sourceString for the first rectangle of enclosed
    // DIAGRAM_MARKER characters at least DIAGRAM_START.length wide
    for (var i = sourceString.indexOf(DIAGRAM_START);
         i >= 0;
         i = sourceString.indexOf(DIAGRAM_START, i + DIAGRAM_START.length)) {

        // We found what looks like a diagram start. See if it has either a full border of
        // aligned '*' characters, or top-left-bottom borders and nothing but white space on
        // the left.
        
        // Look backwards to find the beginning of the line (or of the string)
        // and measure the start character relative to it
        var lineBeginning = max(0, sourceString.lastIndexOf('\n', i)) + 1;
        var xMin = i - lineBeginning;
        
        // Find the first non-diagram character on this line...or the end of the entire source string
        var j;
        for (j = i + DIAGRAM_START.length; sourceString[j] === DIAGRAM_MARKER; ++j) {}
        var xMax = j - lineBeginning - 1;
        
        // We have a potential hit. Start accumulating a result. If there was anything
        // between the newline and the diagram, move it to the after string for proper alignment.
        var result = {
            beforeString: sourceString.ss(0, lineBeginning), 
            diagramString: '',
            alignmentHint: 'center', 
            afterString: sourceString.ss(lineBeginning, i).rp(/[ \t]+$/, ' ')
        };

        var nextLineBeginning = 0, wideCharacters = 0;
        var textOnLeft = false, textOnRight = false;
        var noRightBorder = false;

        advance();
                                  
        // Now, see if the pattern repeats on subsequent lines
        for (var good = true, previousEnding = j; good; ) {
            // Find the next line
            lineBeginning = nextLineBeginning;
            advance();
            if (lineBeginning === 0) {
                // Hit the end of the string before the end of the pattern
                return noDiagramResult; 
            }
            
            if (textOnLeft) {
                // Even if there is text on *both* sides
                result.alignmentHint = 'floatright';
            } else if (textOnRight) {
                result.alignmentHint = 'floatleft';
            }
            
            // See if there are markers at the correct locations on the next line
            if ((sourceString[lineBeginning + xMin] === DIAGRAM_MARKER) && 
                (! textOnLeft || (sourceString[lineBeginning + xMax + wideCharacters] === DIAGRAM_MARKER))) {

                // See if there's a complete line of DIAGRAM_MARKER, which would end the diagram
                var x;
                for (x = xMin; (x < xMax) && (sourceString[lineBeginning + x] === DIAGRAM_MARKER); ++x) {}
           
                var begin = lineBeginning + xMin;
                var end   = lineBeginning + xMax + wideCharacters;
                
                if (! textOnLeft) {
                    // This may be an incomplete line
                    var newlineLocation = sourceString.indexOf('\n', begin);
                    if (newlineLocation !== -1) {
                        end = Math.min(end, newlineLocation);
                    }
                }

                // Trim any excess whitespace caused by our truncation because Markdown will
                // interpret that as fixed-formatted lines
                result.afterString += sourceString.ss(previousEnding, begin).rp(/^[ \t]*[ \t]/, ' ').rp(/[ \t][ \t]*$/, ' ');
                if (x === xMax) {
                    // We found the last row. Put everything else into
                    // the afterString and return the result.
                
                    result.afterString += sourceString.ss(lineBeginning + xMax + 1);
                    return result;
                } else {
                    // A line of a diagram. Extract everything before
                    // the diagram line started into the string of
                    // content to be placed after the diagram in the
                    // final HTML
                    result.diagramString += sourceString.ss(begin + 1, end) + '\n';
                    previousEnding = end + 1;
                }
            } else {
                // Found an incorrectly delimited line. Abort
                // processing of this potential diagram, which is now
                // known to NOT be a diagram after all.
                good = false;
            }
        } // Iterate over verticals in the potential box
    } // Search for the start

    return noDiagramResult;
}

/** 
    Find the specified delimiterRegExp used as a quote (e.g., *foo*)
    and replace it with the HTML tag and optional attributes.
*/
function replaceMatched(string, delimiterRegExp, tag, attribs) {
    var delimiter = delimiterRegExp.source;
    var flanking = '[^ \\t\\n' + delimiter + ']';
    var pattern  = '([^A-Za-z0-9])(' + delimiter + ')' +
        '(' + flanking + '.*?(\\n.+?)*?)' + 
        delimiter + '(?![A-Za-z0-9])';

    return string.rp(new RegExp(pattern, 'g'), 
                          '$1<' + tag + (attribs ? ' ' + attribs : '') +
                          '>$3</' + tag + '>');
}
    
/** Maruku ("github")-style table processing */
function replaceTables(s, protect) {
    var TABLE_ROW       = /(?:\n[ \t]*(?:(?:\|?[ \t\S]+?(?:\|[ \t\S]+?)+\|?)|\|[ \t\S]+\|)(?=\n))/.source;
    var TABLE_SEPARATOR = /\n[ \t]*(?:(?:\|? *\:?-+\:?(?: *\| *\:?-+\:?)+ *\|?|)|\|[\:-]+\|)(?=\n)/.source;
    var TABLE_CAPTION   = /\n[ \t]*\[[^\n\|]+\][ \t]*(?=\n)/.source;
    var TABLE_REGEXP    = new RegExp(TABLE_ROW + TABLE_SEPARATOR + TABLE_ROW + '+(' + TABLE_CAPTION + ')?', 'g');

    function trimTableRowEnds(row) {
        return row.trim().rp(/^\||\|$/g, '');
    }

    s = s.rp(TABLE_REGEXP, function (match) {
        // Found a table, actually parse it by rows
        var rowArray = match.split('\n');
        
        var result = '';
        
        // Skip the bogus leading row
        var startRow = (rowArray[0] === '') ? 1 : 0;

        var caption = rowArray[rowArray.length - 1].trim();

        if ((caption.length > 3) && (caption[0] === '[') && (caption[caption.length - 1] === ']')) {
            // Remove the caption from the row array
            rowArray.pop();
            caption = caption.ss(1, caption.length - 1);
        } else {
            caption = undefined;
        }

        // Parse the separator row for left/center/right-indicating colons
        var columnStyle = [];
        trimTableRowEnds(rowArray[startRow + 1]).rp(/:?-+:?/g, function (match) {
            var left = (match[0] === ':');
            var right = (match[match.length - 1] === ':');
            columnStyle.push(protect(' style="text-align:' + ((left && right) ? 'center' : (right ? 'right' : 'left')) + '"'));
        });

        var row = rowArray[startRow + 1].trim();
        var hasLeadingBar  = row[0] === '|';
        var hasTrailingBar = row[row.length - 1] === '|';
        
        var tag = 'th';
        
        for (var r = startRow; r < rowArray.length; ++r) {
            // Remove leading and trailing whitespace and column delimiters
            row = rowArray[r].trim();
            
            if (! hasLeadingBar && (row[0] === '|')) {
                // Empty first column
                row = '&nbsp;' + row;
            }
            
            if (! hasTrailingBar && (row[row.length - 1] === '|')) {
                // Empty last column
                row += '&nbsp;';
            }
            
            row = trimTableRowEnds(row);
            var i = 0;
            result += entag('tr', '<' + tag + columnStyle[0] + '> ' + 
                            row.rp(/ *\| */g, function () {
                                ++i;
                                return ' </' + tag + '><' + tag + columnStyle[i] + '> ';
                            }) + ' </' + tag + '>') + '\n';
            
            // Skip the header-separator row
            if (r == startRow) { 
                ++r; 
                tag = 'td';
            }
        }
        
        result = entag('table', result, protect('class="table"'));

        if (caption) {
            caption = entag('center', entag('div', caption, protect('class="tablecaption"')));
            if (option('captionAbove', 'table')) {
                result = caption + result;
            } else {
                result = '\n' + result + caption;
            }
        }

        return entag('div', result, "class='table'");
    });

    return s;
}


function replaceLists(s, protect) {
    // Identify task list bullets in a few patterns and reformat them to a standard format for
    // easier processing.
    s = s.rp(/^(\s*)(?:-\s*)?(?:\[ \]|\u2610)(\s+)/mg, '$1\u2610$2');
    s = s.rp(/^(\s*)(?:-\s*)?(?:\[[xX]\]|\u2611)(\s+)/mg, '$1\u2611$2');
        
    // Identify list blocks:
    // Blank line or line ending in colon, line that starts with #., *, +, -, ☑, or ☐
    // and then any number of lines until another blank line
    var BLANK_LINES = /\n\s*\n/.source;

    // Preceding line ending in a colon

    // \u2610 is the ballot box (unchecked box) character
    var PREFIX     = /[:,]\s*\n/.source;
    var LIST_BLOCK_REGEXP = 
        new RegExp('(' + PREFIX + '|' + BLANK_LINES + '|<p>\s*\n|<br/>\s*\n?)' +
                   /((?:[ \t]*(?:\d+\.|-|\+|\*|\u2611|\u2610)(?:[ \t]+.+\n(?:[ \t]*\n)?)+)+)/.source, 'gm');

    var keepGoing = true;

    var ATTRIBS = {'+': protect('class="plus"'), '-': protect('class="minus"'), '*': protect('class="asterisk"'),
                   '\u2611': protect('class="checked"'), '\u2610': protect('class="unchecked"')};
    var NUMBER_ATTRIBS = protect('class="number"');

    // Sometimes the list regexp grabs too much because subsequent lines are indented *less*
    // than the first line. So, if that case is found, re-run the regexp.
    while (keepGoing) {
        keepGoing = false;
        s = s.rp(LIST_BLOCK_REGEXP, function (match, prefix, block) {
            var result = prefix;
            
            // Contains {indentLevel, tag}
            var stack = [];
            var current = {indentLevel: -1};
            
            /* function logStack(stack) {
               var s = '[';
               stack.forEach(function(v) { s += v.indentLevel + ', '; });
               console.log(s.ss(0, s.length - 2) + ']');
               } */
            block.split('\n').forEach(function (line) {
                var trimmed     = line.rp(/^\s*/, '');
                
                var indentLevel = line.length - trimmed.length;
                
                // Add a CSS class based on the type of list bullet
                var attribs = ATTRIBS[trimmed[0]];
                var isUnordered = !! attribs; // JavaScript for: attribs !== undefined
                attribs = attribs || NUMBER_ATTRIBS;
                var isOrdered   = /^\d+\.[ \t]/.test(trimmed);
                var isBlank     = trimmed === '';
                var start       = isOrdered ? ' ' + protect('start=' + trimmed.match(/^\d+/)[0]) : '';

                if (isOrdered || isUnordered) {
                    // Add the indentation for the bullet itself
                    indentLevel += 2;
                }

                if (! current) {
                    // Went below top-level indent
                    result += '\n' + line;
                } else if (! isOrdered && ! isUnordered && (isBlank || (indentLevel >= current.indentLevel))) {
                    // Line without a marker
                    result += '\n' + current.indentChars + line;
                } else {
                    //console.log(indentLevel + ":" + line);
                    if (indentLevel !== current.indentLevel) {
                        // Enter or leave indentation level
                        if ((current.indentLevel !== -1) && (indentLevel < current.indentLevel)) {
                            while (current && (indentLevel < current.indentLevel)) {
                                stack.pop();
                                // End the current list and decrease indentation
                                result += '\n</li></' + current.tag + '>';
                                current = stack[stack.length - 1];
                            }
                        } else {
                            // Start a new list that is more indented
                            current = {indentLevel: indentLevel,
                                       tag:         isOrdered ? 'ol' : 'ul',
                                       // Subtract off the two indent characters we added above
                                       indentChars: line.ss(0, indentLevel - 2)};
                            stack.push(current);
                            result += '\n<' + current.tag + start + '>';
                        }
                    } else if (current.indentLevel !== -1) {
                        // End previous list item, if there was one
                        result += '\n</li>';
                    } // Indent level changed
                    
                    if (current) {
                        // Add the list item
                        result += '\n' + current.indentChars + '<li ' + attribs + '>' + trimmed.rp(/^(\d+\.|-|\+|\*|\u2611|\u2610) /, '');
                    } else {
                        // Just reached something that is *less* indented than the root--
                        // copy forward and then re-process that list
                        result += '\n' + line;
                        keepGoing = true;
                    }
                }
            }); // For each line

            // Remove trailing whitespace
            result = result.replace(/\s+$/,'');
            
            // Finish the last item and anything else on the stack (if needed)
            for (current = stack.pop(); current; current = stack.pop()) {
                result += '</li></' + current.tag + '>';
            }
       
            return result + '\n\n';
        });
    } // while keep going

    return s;
}


/** 
    Identifies schedule lists, which look like:

  date: title
    events

  Where date must contain a day, month, and four-number year and may
  also contain a day of the week.  Note that the date must not be
  indented and the events must be indented.

  Multiple events per date are permitted.
*/
function replaceScheduleLists(str, protect) {
    // Must open with something other than indentation or a list
    // marker.  There must be a four-digit number somewhere on the
    // line. Exclude lines that begin with an HTML tag...this will
    // avoid parsing headers that have dates in them.
    var BEGINNING = /^(?:[^\|<>\s-\+\*\d].*[12]\d{3}(?!\d).*?|(?:[12]\d{3}(?!\.).*\d.*?)|(?:\d{1,3}(?!\.).*[12]\d{3}(?!\d).*?))/.source;

    // There must be at least one more number in a date, a colon, and then some more text
    var DATE_AND_TITLE = '(' + BEGINNING + '):' + /[ \t]+([^ \t\n].*)\n/.source;

    // The body of the schedule item. It may begin with a blank line and contain
    // multiple paragraphs separated by blank lines...as long as there is indenting
    var EVENTS = /(?:[ \t]*\n)?((?:[ \t]+.+\n(?:[ \t]*\n){0,3})*)/.source;
    var ENTRY = DATE_AND_TITLE + EVENTS;

    var BLANK_LINE = '\n[ \t]*\n';
    var ENTRY_REGEXP = new RegExp(ENTRY, 'gm');

    var rowAttribs = protect('valign="top"');
    var dateTDAttribs = protect('style="width:100px;padding-right:15px" rowspan="2"');
    var eventTDAttribs = protect('style="padding-bottom:25px"');

    var DAY_NAME   = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'].map(keyword);
    var MONTH_NAME = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].map(keyword);
    var MONTH_FULL_NAME = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].map(keyword);

    function clean(s) { return s.toLowerCase().rp('.', ''); }
    var LOWERCASE_MONTH_NAME = MONTH_NAME.map(clean);
    var LOWERCASE_MONTH_FULL_NAME = MONTH_FULL_NAME.map(clean);

    // Allow a period (and capture it) after each word, but eliminate
    // the periods that are in abbreviations so that they do not appear
    // in the regexp as wildcards or word breaks
    var MONTH_NAME_LIST = '\\b' + MONTH_NAME.concat(MONTH_FULL_NAME).join('(?:\\.|\\b)|\\b').rp(/([^\\])\./g, '$1') + '(?:\\.|\\b)';

    // Used to mark the center of each day. Not close to midnight to avoid daylight
    // savings problems.
    var standardHour = 9;

    try {
        var scheduleNumber = 0;
        str = str.rp(new RegExp(BLANK_LINE + '(' + ENTRY + '){2,}', 'gm'),
                     function (schedule) {
                       ++scheduleNumber;
                       // Each entry has the form {date:date, title:string, text:string}
                       var entryArray = [];

                       // Now parse the schedule into individual day entries

                       var anyWeekendEvents = false;

                       schedule.rp
                         (ENTRY_REGEXP,
                          function (entry, date, title, events) {
                              // Remove the day from the date (we'll reconstruct it below). This is actually unnecessary, since we
                              // explicitly compute the value anyway and the parser is robust to extra characters, but it aides
                              // in debugging.
                              // 
                              // date = date.rp(/(?:(?:sun|mon|tues|wednes|thurs|fri|satur)day|(?:sun|mon|tue|wed|thu|fri|sat)\.?|(?:su|mo|tu|we|th|fr|sa)),?/gi, '');
                              
                              // Parse the date. The Javascript Date class's parser is useless because it
                              // is locale dependent, so we do this with a regexp.
                              
                              var year = '', month = '', day = '', parenthesized = false;
                              
                              date = date.trim();
                              
                              if ((date[0] === '(') && (date.slice(-1) === ')')) {
                                  // This is a parenthesized entry
                                  date = date.slice(1, -1);
                                  parenthesized = true;
                              }
                              
                              // DD MONTH YYYY
                              var DD_MONTH_YYYY = RegExp('([0123]?\\d)\\D+([01]?\\d|' + MONTH_NAME_LIST + ')\\D+([12]\\d{3})', 'i')
                              var match = date.match(DD_MONTH_YYYY);
                              
                              if (match) {
                                  day = match[1]; month = match[2]; year = match[3];
                              } else {
                                  // YYYY MONTH DD
                                  match = date.match(RegExp('([12]\\d{3})\\D+([01]?\\d|' + MONTH_NAME_LIST + ')\\D+([0123]?\\d)', 'i')); 
                                  if (match) {
                                      day = match[3]; month = match[2]; year = match[1];
                                  } else {
                                      // MONTH DD YYYY
                                      match = date.match(RegExp('(' + MONTH_NAME_LIST + ')\\D+([0123]?\\d)\\D+([12]\\d{3})', 'i'));
                                      if (match) {
                                          day = match[2]; month = match[1]; year = match[3];
                                      } else {
                                          throw "Could not parse date";
                                      }
                                  }
                              }
                              
                              // Reconstruct standardized date format
                              date = day + '&nbsp;' + keyword(month) + '&nbsp;' + year;

                              // Detect the month
                              var monthNumber = parseInt(month) - 1;
                              if (isNaN(monthNumber)) {
                                  var target = clean(month);
                                  monthNumber = LOWERCASE_MONTH_NAME.indexOf(target);
                                  if (monthNumber === -1) {
                                      monthNumber = LOWERCASE_MONTH_FULL_NAME.indexOf(target);
                                  }
                              }

                              var dateVal = new Date(Date.UTC(parseInt(year), monthNumber, parseInt(day), standardHour));
                              // Reconstruct the day of the week
                              var dayOfWeek = dateVal.getUTCDay();
                              date = DAY_NAME[dayOfWeek] + '<br/>' + date;
                              
                              anyWeekendEvents = anyWeekendEvents || (dayOfWeek === 0) || (dayOfWeek === 6);
                              
                              entryArray.push({date: dateVal, 
                                               title: title,
                                               sourceOrder: entryArray.length,
                                               parenthesized: parenthesized,
                                               
                                               // Don't show text if parenthesized with no body
                                               text: parenthesized ? '' :
                                               entag('tr',
                                                     entag('td', 
                                                           '<a ' + protect('class="target" name="schedule' + scheduleNumber + '_' + dateVal.getUTCFullYear() + '-' + (dateVal.getUTCMonth() + 1) + '-' + dateVal.getUTCDate() + '"') + '>&nbsp;</a>' +
                                                           date, dateTDAttribs) + 
                                                     entag('td', entag('b', title)), rowAttribs) + 
                                               entag('tr', entag('td', '\n\n' + events, eventTDAttribs), rowAttribs)});
                              
                              return '';
                          });
                         
                         // Shallow copy the entries to bypass sorting if needed
                         var sourceEntryArray = option('sortScheduleLists') ? entryArray : entryArray.slice(0);

                         // Sort by date
                         entryArray.sort(function (a, b) {
                             // Javascript's sort is not specified to be
                             // stable, so we have to preserve
                             // sourceOrder in ties.
                             var ta = a.date.getTime();
                             var tb = b.date.getTime();
                             return (ta === tb) ? (a.sourceOrder - b.sourceOrder) : (ta - tb);
                         });
                         
                         var MILLISECONDS_PER_DAY = 1000 * 60 * 60 * 24;
                         
                         // May be slightly off due to daylight savings time
                         var approximateDaySpan = (entryArray[entryArray.length - 1].date.getTime() - entryArray[0].date.getTime()) / MILLISECONDS_PER_DAY;
                         
                         var today = new Date();
                         // Move back to midnight
                         today = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate(), standardHour));
                         
                         var calendar = '';
                         // Make a calendar view with links, if suitable
                         if ((approximateDaySpan > 14) && (approximateDaySpan / entryArray.length < 16)) {
                             var DAY_HEADER_ATTRIBS = protect('colspan="2" width="14%" style="padding-top:5px;text-align:center;font-style:italic"');
                             var DATE_ATTRIBS       = protect('width="1%" height="30px" style="text-align:right;border:1px solid #EEE;border-right:none;"');
                             var FADED_ATTRIBS      = protect('width="1%" height="30px" style="color:#BBB;text-align:right;"');
                             var ENTRY_ATTRIBS      = protect('width="14%" style="border:1px solid #EEE;border-left:none;"');
                             var PARENTHESIZED_ATTRIBS = protect('class="parenthesized"');
                             
                             // Find the first day of the first month
                             var date = entryArray[0].date;
                             var index = 0;
                             
                             var hideWeekends = ! anyWeekendEvents && option('hideEmptyWeekends');
                             var showDate = hideWeekends ? function(date) { return (date.getUTCDay() > 0) && (date.getUTCDay() < 6);} : function() { return true; };
                             
                             var sameDay = function (d1, d2) {
                                 // Account for daylight savings time
                                 return (abs(d1.getTime() - d2.getTime()) < MILLISECONDS_PER_DAY / 2);
                             }
                             
                             // Go to the first of the month
                             date = new Date(date.getUTCFullYear(), date.getUTCMonth(), 1, standardHour);
                             
                             while (date.getTime() < entryArray[entryArray.length - 1].date.getTime()) {
                                 
                                 // Create the calendar header
                                 calendar += '<table ' + protect('class="calendar"') + '>\n' +
                                     entag('tr', entag('th', MONTH_FULL_NAME[date.getUTCMonth()] + ' ' + date.getUTCFullYear(), protect('colspan="14"'))) + '<tr>';
                                 
                                 (hideWeekends ? DAY_NAME.slice(1, 6) : DAY_NAME).forEach(function (name) {
                                     calendar += entag('td', name, DAY_HEADER_ATTRIBS);
                                 });
                                 calendar += '</tr>';
                                 
                                 // Go back into the previous month to reach a Sunday. Check the time at noon
                                 // to avoid problems with daylight saving time occuring early in the morning
                                 while (date.getUTCDay() !== 0) { 
                                     date = new Date(date.getTime() - MILLISECONDS_PER_DAY); 
                                 }
                                 
                                 // Insert the days from the previous month
                                 if (date.getDate() !== 1) {
                                     calendar += '<tr ' + rowAttribs + '>';
                                     while (date.getDate() !== 1) {
                                         if (showDate(date)) { calendar += '<td ' + FADED_ATTRIBS + '>' + date.getUTCDate() + '</td><td>&nbsp;</td>'; }
                                         date = new Date(date.getTime() + MILLISECONDS_PER_DAY);
                                     }
                                 }
                                 
                                 // Run until the end of the month
                                 do {
                                     if (date.getUTCDay() === 0) {
                                         // Sunday, start a row
                                         calendar += '<tr ' + rowAttribs + '>';
                                     }
                                     
                                     if (showDate(date)) {
                                         var attribs = '';
                                         if (sameDay(date, today)) {
                                             attribs = protect('class="today"');
                                         }
                                         
                                         // Insert links as needed from entries
                                         var contents = '';
                                         
                                         for (var entry = entryArray[index]; entry && sameDay(entry.date, date); ++index, entry = entryArray[index]) {
                                             if (contents) { contents += '<br/>'; }
                                             if (entry.parenthesized) {
                                                 // Parenthesized with no body, no need for a link
                                                 contents += entag('span', entry.title, PARENTHESIZED_ATTRIBS);
                                             } else {
                                                 contents += entag('a', entry.title, protect('href="#schedule' + scheduleNumber + '_' + date.getUTCFullYear() + '-' + (date.getUTCMonth() + 1) + '-' + date.getUTCDate() + '"'));
                                             }
                                         }
                                         
                                         if (contents) {
                                             calendar += entag('td', entag('b', date.getUTCDate()), DATE_ATTRIBS + attribs) + entag('td', contents, ENTRY_ATTRIBS + attribs);
                                         } else {
                                             calendar += '<td ' + DATE_ATTRIBS + attribs + '></a>' + date.getUTCDate() + '</td><td ' + ENTRY_ATTRIBS + attribs + '> &nbsp; </td>';
                                         }
                                     }                                   
                                     
                                     if (date.getUTCDay() === 6) {
                                         // Saturday, end a row
                                         calendar += '</tr>';
                                     }
                                     
                                     // Go to (approximately) the next day
                                     date = new Date(date.getTime() + MILLISECONDS_PER_DAY);
                                 } while (date.getUTCDate() > 1);
                                 
                               // Finish out the week after the end of the month
                               if (date.getUTCDay() !== 0) {
                                   while (date.getUTCDay() !== 0) {
                                       if (showDate(date)) { calendar += '<td ' + FADED_ATTRIBS + '>' + date.getUTCDate() + '</td><td>&nbsp</td>'; }
                                       date = new Date(date.getTime() + MILLISECONDS_PER_DAY);
                                   }
                                   
                                   calendar += '</tr>';
                               }

                               calendar += '</table><br/>\n';

                               // Go to the first of the (new) month
                               date = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1, standardHour));
                               
                           } // Until all days covered
                       } // if add calendar

                       // Construct the schedule
                       schedule = '';
                       sourceEntryArray.forEach(function (entry) {
                           schedule += entry.text;
                       });

                       return '\n\n' + calendar + entag('table', schedule, protect('class="schedule"')) + '\n\n';
                     });
    } catch (ignore) {
        // Maybe this wasn't a schedule after all, since we couldn't parse a date. Don't alarm
        // the user, though
    }

    return str;
}


/**
 Term
 :     description, which might be multiple 
       lines and include blanks.

 Next Term

becomes

<dl>
  <dt>Term</dt>
  <dd> description, which might be multiple 
       lines and include blanks.</dd>
  <dt>Next Term</dt>
</dl>

... unless it is very short, in which case it becomes a table.

*/
function replaceDefinitionLists(s, protect) {
    var TERM       = /^.+\n:(?=[ \t])/.source;

    // Definition can contain multiple paragraphs
    var DEFINITION = '(\s*\n|[: \t].+\n)+';

    s = s.rp(new RegExp('(' + TERM + DEFINITION + ')+', 'gm'),
             function (block) {
                 
                 var list = [];

                 // Parse the block
                 var currentEntry = null;
 
                 block.split('\n').forEach(function (line, i) {
                     // What kind of line is this?
                     if (line.trim().length === 0) {
                         if (currentEntry) {
                             // Empty line
                             currentEntry.definition += '\n';
                         }
                     } else if (! /\s/.test(line[0]) && (line[0] !== ':')) {
                         currentEntry = {term: line, definition: ''};
                         list.push(currentEntry);
                     } else {
                         // Add the line to the current definition, stripping any single leading ':'
                         if (line[0] === ':') { line = ' ' + line.ss(1); }
                         currentEntry.definition += line + '\n';
                     }
                 });

                 var longestDefinition = 0;
                 list.forEach(function (entry) {
                     if (/\n\s*\n/.test(entry.definition.trim())) {
                         // This definition contains multiple paragraphs. Force it into long mode
                         longestDefinition = Infinity;
                     } else {
                         // Normal case
                         longestDefinition = max(longestDefinition, unescapeHTMLEntities(removeHTMLTags(entry.definition)).length);
                     }
                 });

                 var result = '';
                 var definitionStyle = option('definitionStyle');
                 if ((definitionStyle === 'short') || ((definitionStyle !== 'long') && (longestDefinition < 160))) {
                     var rowAttribs = protect('valign=top');
                     // This list has short definitions. Format it as a table
                     list.forEach(function (entry) {
                         result += entag('tr',
                                         entag('td', entag('dt', entry.term)) + 
                                         entag('td', entag('dd', entag('p', entry.definition))), 
                                         rowAttribs);
                     });
                     result = entag('table', result);

                 } else {
                     list.forEach(function (entry) {
                         // Leave *two* blanks at the start of a
                         // definition so that subsequent processing
                         // can detect block formatting within it.
                         result += entag('dt', entry.term) + entag('dd', entag('p', entry.definition));
                     });
                 }

                 return entag('dl', result);

             });

    return s;
}


/** Inserts a table of contents in the document and then returns
    [string, table], where the table maps strings to levels. */
function insertTableOfContents(s, protect, exposer) {

    // Gather headers for table of contents (TOC). We
    // accumulate a long and short TOC and then choose which
    // to insert at the end.
    var fullTOC = '<a href="#" class="tocTop">(Top)</a><br/>\n';
    var shortTOC = '';

    // names of parent sections
    var nameStack = [];
    
    // headerCounter[i] is the current counter for header level (i - 1)
    var headerCounter = [0];
    var currentLevel = 0;
    var numAboveLevel1 = 0;

    var table = {};
    s = s.rp(/<h([1-6])>(.*?)<\/h\1>/gi, function (header, level, text) {
        level = parseInt(level)
        text = text.trim();
        
        // If becoming more nested:
        for (var i = currentLevel; i < level; ++i) {
            nameStack[i] = '';
            headerCounter[i] = 0;
        }
        
        // If becoming less nested:
        headerCounter.splice(level, currentLevel - level);
        nameStack.splice(level, currentLevel - level);
        currentLevel = level;

        ++headerCounter[currentLevel - 1];
        
        // Generate a unique name for this element
        var number = headerCounter.join('.');

        // legacy, for when toc links were based on
        // numbers instead of mangled names
        var oldname = 'toc' + number;

        var cleanText = removeHTMLTags(exposer(text)).trim().toLowerCase();
        
        table[cleanText] = number;

        // Remove links from the title itself
        text = text.rp(/<a\s.*>(.*?)<\/a>/g, '$1');

        nameStack[currentLevel - 1] = mangle(cleanText);

        var name = nameStack.join('/');

        // Only insert for the first three levels
        if (level <= 3) {
            // Indent and append (the Array() call generates spaces)
            fullTOC += Array(level).join('&nbsp;&nbsp;') + '<a href="#' + name + '" class="level' + level + '"><span class="tocNumber">' + number + '&nbsp; </span>' + text + '</a><br/>\n';
            
            if (level === 1) {
                shortTOC += ' &middot; <a href="#' + name + '">' + text + '</a>';
            } else {
                ++numAboveLevel1;
            }
        }

        return entag('a', '&nbsp;', protect('class="target" name="' + name + '"')) +
            entag('a', '&nbsp;', protect('class="target" name="' + oldname + '"')) +
            header;
    });

    if (shortTOC.length > 0) {
        // Strip the leading " &middot; "
        shortTOC = shortTOC.ss(10);
    }
    
    var numLevel1 = headerCounter[0];
    var numHeaders = numLevel1 + numAboveLevel1;

    // The location of the first header is indicative of the length of
    // the abstract...as well as where we insert. The first header may be accompanied by
    // <a name> tags, which we want to appear before.
    var firstHeaderLocation = s.regexIndexOf(/((<a\s+\S+>&nbsp;<\/a>)\s*)*?<h\d>/i);
    if (firstHeaderLocation === -1) { firstHeaderLocation = 0; }

    var AFTER_TITLES = '<div class="afterTitles"><\/div>';
    var insertLocation = s.indexOf(AFTER_TITLES);
    if (insertLocation === -1) {
        insertLocation = 0;
    } else {
        insertLocation += AFTER_TITLES.length;
    }

    // Which TOC style should we use?
    var tocStyle = option('tocStyle');

    var TOC = '';
    if ((tocStyle === 'auto') || (tocStyle === '')) {
        if (((numHeaders < 4) && (numLevel1 <= 1)) || (s.length < 2048)) {
            // No TOC; this document is really short
            tocStyle = 'none';
        } else if ((numLevel1 < 7) && (numHeaders / numLevel1 < 2.5)) {
            // We can use the short TOC
            tocStyle = 'short';
        } else if ((firstHeaderLocation === -1) || (firstHeaderLocation / 55 > numHeaders)) {
            // The abstract is long enough to float alongside, and there
            // are not too many levels.        
            // Insert the medium-length TOC floating
            tocStyle = 'medium';
        } else {
            // This is a long table of contents or a short abstract
            // Insert a long toc...right before the first header
            tocStyle = 'long';
        }
    }

    switch (tocStyle) {
    case 'none':
    case '':
        break;

    case 'short':
        TOC = '<div class="shortTOC">' + shortTOC + '</div>';
        break;

    case 'medium':
        TOC = '<div class="mediumTOC"><center><b>' + keyword('Contents') + '</b></center><p>' + fullTOC + '</p></div>';
        break;

    case 'long':
        insertLocation = firstHeaderLocation;
        TOC = '<div class="longTOC"><div class="tocHeader">' + keyword('Contents') + '</div><p>' + fullTOC + '</p></div>';
        break;

    default:
        console.log('markdeepOptions.tocStyle = "' + tocStyle + '" specified in your document is not a legal value');
    }

    s = s.ss(0, insertLocation) + TOC + s.ss(insertLocation);

    return [s, table];
}


function escapeRegExpCharacters(str) {
    return str.rp(/([\.\[\]\(\)\*\+\?\^\$\\\{\}\|])/g, '\\$1');
}


/** Returns true if there are at least two newlines in each of the arguments */
function isolated(preSpaces, postSpaces) {
    if (preSpaces && postSpaces) {
        preSpaces  = preSpaces.match(/\n/g);
        postSpaces = postSpaces.match(/\n/g);
        return preSpaces && (preSpaces.length > 1) && postSpaces && (postSpaces.length > 1);
    } else {
        return false;
    }
}


/**
    Performs Markdeep processing on str, which must be a string or a
    DOM element.  Returns a string that is the HTML to display for the
    body. The result does not include the header: Markdeep stylesheet
    and script tags for including a math library, or the Markdeep
    signature footer.

    Optional argument elementMode defaults to true. This avoids turning a bold first word into a 
    title or introducing a table of contents. Section captions are unaffected by this argument.
    Set elementMode = false if processing a whole document instead of an internal node.

 */
function markdeepToHTML(str, elementMode) {
    // Map names to the number used for end notes, in the order
    // encountered in the text.
    var endNoteTable = {}, endNoteCount = 0;

    // Reference links
    var referenceLinkTable = {};

    // In the private use area
    var PROTECT_CHARACTER = '\ue010';

    // Use base 32 for encoding numbers, which is efficient in terms of 
    // characters but avoids 'x' to avoid the pattern \dx\d, which Markdeep would
    // beautify as a dimension
    var PROTECT_RADIX     = 32;
    var protectedStringArray = [];

    // Gives 1e6 possible sequences in base 32, which should be sufficient
    var PROTECT_DIGITS    = 4;

    // Put the protect character at BOTH ends to avoid having the protected number encoding
    // look like an actual number to further markdown processing
    var PROTECT_REGEXP    = RegExp(PROTECT_CHARACTER + '[0-9a-w]{' + PROTECT_DIGITS + ',' + PROTECT_DIGITS + '}' + PROTECT_CHARACTER, 'g');

    /** Given an arbitrary string, returns an escaped identifier
        string to temporarily replace it with to prevent Markdeep from
        processing the contents. See expose() */
    function protect(s) {
        // Generate the replacement index, converted to an alphanumeric string
        var i = (protectedStringArray.push(s) - 1).toString(PROTECT_RADIX);

        // Ensure fixed length
        while (i.length < PROTECT_DIGITS) {
            i = '0' + i;
        }

        return PROTECT_CHARACTER + i + PROTECT_CHARACTER;
    }

    var exposeRan = false;
    /** Given the escaped identifier string from protect(), returns
        the orginal string. */
    function expose(i) {
        // Strip the escape character and parse, then look up in the
        // dictionary.
        var j = parseInt(i.ss(1, i.length - 1), PROTECT_RADIX);
        exposeRan = true;
        return protectedStringArray[j];
    }

    /** First-class function to pass to String.replace to protect a
        sequence defined by a regular expression. */
    function protector(match, protectee) {
        return protect(protectee);
    }

    function protectorWithPrefix(match, prefix, protectee) {
        return prefix + protect(protectee);
    }

    // SECTION HEADERS
    // This is common code for numbered headers. No-number ATX headers are processed
    // separately
    function makeHeaderFunc(level) {
        return function (match, header) {
            return '\n\n</p>\n<a ' + protect('class="target" name="' + mangle(removeHTMLTags(header.rp(PROTECT_REGEXP, expose))) + '"') + 
                '>&nbsp;</a>' + entag('h' + level, header) + '\n<p>\n\n';
        }
    }

    if (elementMode === undefined) { 
        elementMode = true;
    }
    
    if (str.innerHTML !== undefined) {
        str = str.innerHTML;
    }

    // Prefix a newline so that blocks beginning at the top of the
    // document are processed correctly
    str = '\n\n' + str;

    // Replace pre-formatted script tags that are used to protect
    // less-than signs, e.g., in std::vector<Value>
    str = str.rp(/<script\s+type\s*=\s*['"]preformatted['"]\s*>([\s\S]*?)<\/script>/gi, '$1');

    function replaceDiagrams(str) {
        var result = extractDiagram(str);
        if (result.diagramString) {
            var CAPTION_REGEXP = /^\n*[ \t]*\[[^\n]+\][ \t]*(?=\n)/;
            result.afterString = result.afterString.rp(CAPTION_REGEXP, function (caption) {
                // Strip whitespace and enclosing brackets from the caption
                caption = caption.trim();
                caption = caption.ss(1, caption.length - 1);
                
                result.caption = entag('center', entag('div', caption, protect('class="imagecaption"')));
                return '';
            });

            var diagramSVG = diagramToSVG(result.diagramString, result.alignmentHint);
            var captionAbove = option('captionAbove', 'diagram')

            return result.beforeString +
                (result.caption && captionAbove ? result.caption : '') +
                diagramSVG +
                (result.caption && ! captionAbove ? result.caption : '') + '\n' +
                replaceDiagrams(result.afterString);
        } else {
            return str;
        }
    }

    // CODE FENCES, with styles. Do this before other processing so that their code is
    // protected from further Markdown processing
    var stylizeFence = function (cssClass, symbol) {
        var pattern = new RegExp('\n([ \\t]*)' + symbol + '{3,}([ \\t]*\\S*)([ \\t]+.+)?\n([\\s\\S]+?)\n\\1' + symbol + '{3,}[ \t]*\n([ \\t]*\\[.+(?:\n.+){0,3}\\])?', 'g');
        
        str = str.rp(pattern, function(match, indent, lang, cssSubClass, sourceCode, caption) {
            if (caption) {
                caption = caption.trim();
                caption = entag('center', '<div ' + protect('class="listingcaption ' + cssClass + '"') + '>' + caption.ss(1, caption.length - 1) + '</div>') + '\n';
            }
            // Remove the block's own indentation from each line of sourceCode
            sourceCode = sourceCode.rp(new RegExp('(^|\n)' + indent, 'g'), '$1');

            var captionAbove = option('captionAbove', 'listing')
            var nextSourceCode, nextLang, nextCssSubClass;
            var body = [];

            // Process multiple-listing blocks
            do {
                nextSourceCode = nextLang = nextCssSubClass = undefined;
                sourceCode = sourceCode.rp(new RegExp('\\n([ \\t]*)' + symbol + '{3,}([ \\t]*\\S+)([ \\t]+.+)?\n([\\s\\S]*)'),
                                           function (match, indent, lang, cssSubClass, everythingElse) {
                                               nextLang = lang;
                                               nextCssSubClass = cssSubClass;
                                               nextSourceCode = everythingElse;
                                               return '';
                                           });

                // Highlight and append this block
                lang = lang ? lang.trim() : undefined;
                var result;
                if (lang === 'none') {
                    result = hljs.highlightAuto(sourceCode, []);
                } else if (lang === undefined) {
                    result = hljs.highlightAuto(sourceCode);
                } else {
                    try {
                        result = hljs.highlight(lang, sourceCode, true);
                    } catch (e) {
                        // Some unknown language specified. Force to no formatting.
                        result = hljs.highlightAuto(sourceCode, []);
                    }
                }
                
                var highlighted = result.value;

                // Mark each line as a span to support line numbers
                highlighted = highlighted.rp(/^(.*)$/gm, entag('span', '', 'class="line"') + '$1');

                if (cssSubClass) {
                    highlighted = entag('div', highlighted, 'class="' + cssSubClass + '"');
                }

                body.push(highlighted);

                // Advance the next nested block
                sourceCode = nextSourceCode;
                lang = nextLang;
                cssSubClass = nextCssSubClass;
            } while (sourceCode);

            // Insert paragraph close/open tags, since browsers force them anyway around pre tags
            // We need the indent in case this is a code block inside a list that is indented.
            return '\n' + indent + '</p>' + (caption && captionAbove ? caption : '') +
                protect(entag('pre', entag('code', body.join('')), 'class="listing ' + cssClass + '"')) +
                (caption && ! captionAbove ? caption : '') + '<p>\n';
        });
    };

    stylizeFence('tilde', '~');
    stylizeFence('backtick', '`');
    
    // Highlight explicit inline code
    str = str.rp(/<code\s+lang\s*=\s*["']?([^"'\)\[\]\n]+)["'?]\s*>(.*)<\/code>/gi, function (match, lang, body) {
        return entag('code', hljs.highlight(lang, body, true).value, 'lang=' + lang);
    });
    
    // Protect raw <CODE> content
    str = str.rp(/(<code\b.*?<\/code>)/gi, protector);

    // Remove XML/HTML COMMENTS
    // https://html.spec.whatwg.org/multipage/syntax.html#comments
    str = str.rp(/<!--((?!->|>)[\s\S]*?)-->/g, '');

    str = replaceDiagrams(str);
    
    // Protect SVG blocks (including the ones we just inserted)
    str = str.rp(/<svg( .*?)?>([\s\S]*?)<\/svg>/gi, function (match, attribs, body) {
        return '<svg' + protect(attribs) + '>' + protect(body) + '</svg>';
    });
    
    // Protect STYLE blocks
    str = str.rp(/<style>([\s\S]*?)<\/style>/gi, function (match, body) {
        return entag('style', protect(body));
    });

    // Protect the very special case of img tags with newlines and
    // breaks in them AND mismatched angle brackets. This happens for
    // Gravizo graphs.
    str = str.rp(/<img\s+src=(["'])[\s\S]*?\1\s*>/gi, function (match, quote) {
        // Strip the "<img " and ">", and then protect the interior:
        return "<img " + protect(match.ss(5, match.length - 1)) + ">";
    });

    // INLINE CODE: Surrounded in (non-escaped!) back ticks on a single line.  Do this before any other
    // processing except for diagrams to protect code blocks from further interference. Don't process back ticks
    // inside of code fences. Allow a single newline, but not wrapping further because that
    // might just pick up quotes used as other punctuation across lines. Explicitly exclude
    // cases where the second quote immediately preceeds a number, e.g., "the old `97"
    var inlineLang = option('inlineCodeLang');
    var inlineCodeRegexp = /(^|[^\\])`(.*?(?:\n.*?)?[^\n\\`])`(?!\d)/g;
    if (inlineLang) {
        // Syntax highlight as well as converting to code. Protect
        // so that the hljs output isn't itself escaped below.
        var filenameRegexp = /^[a-zA-Z]:\\|^\/[a-zA-Z_\.]|^[a-z]{3,5}:\/\//;
        str = str.rp(inlineCodeRegexp, function (match, before, body) {
            if (filenameRegexp.test(body)) {
                // This looks like a filename, don't highlight it
                return before + entag('code', body);
            } else {
                return before + protect(entag('code', hljs.highlight(inlineLang, body, true).value));
            }
        });
    } else {
        str = str.rp(inlineCodeRegexp, '$1' + entag('code', '$2'));
    }

    // Unescape escaped backticks
    str = str.rp(/\\`/g, '`');
    
    // CODE: Escape angle brackets inside code blocks (including the ones we just introduced),
    // and then protect the blocks themselves
    str = str.rp(/(<code(?: .*?)?>)([\s\S]*?)<\/code>/gi, function (match, open, inlineCode) {
        return protect(open + escapeHTMLEntities(inlineCode) + '</code>');
    });
    
    // PRE: Protect pre blocks
    str = str.rp(/(<pre\b[\s\S]*?<\/pre>)/gi, protector);
    
    // Protect raw HTML attributes from processing
    str = str.rp(/(<\w[^ \n<>]*?[ \t]+)(.*?)(?=\/?>)/g, protectorWithPrefix);

    // End of processing literal blocks
    /////////////////////////////////////////////////////////////////////////////

    // Temporarily hide $$ MathJax LaTeX blocks from Markdown processing (this must
    // come before single $ block detection below)
    str = str.rp(/(\$\$[\s\S]+?\$\$)/g, protector);

    // Convert LaTeX $ ... $ to MathJax, but verify that this
    // actually looks like math and not just dollar
    // signs. Don't rp double-dollar signs. Do this only
    // outside of protected blocks.

    // Also allow LaTeX of the form $...$ if the close tag is not US$ or Can$
    // and there are spaces outside of the dollar signs.
    //
    // Test: " $3 or US$2 and 3$, $x$ $y + \n 2x$ or ($z$) $k$. or $2 or $2".match(pattern) = 
    // ["$x$", "$y +  2x$", "$z$", "$k$"];
    str = str.rp(/((?:[^\w\d]))\$(\S(?:[^\$]*?\S(?!US|Can))??)\$(?![\w\d])/g, '$1\\($2\\)');

    //
    // Literally: find a non-dollar sign, non-number followed
    // by a dollar sign and a space.  Then, find any number of
    // characters until the same pattern reversed, allowing
    // one punctuation character before the final space. We're
    // trying to exclude things like Canadian 1$ and US $1
    // triggering math mode.

    str = str.rp(/((?:[^\w\d]))\$([ \t][^\$]+?[ \t])\$(?![\w\d])/g, '$1\\($2\\)');

    // Temporarily hide MathJax LaTeX blocks from Markdown processing
    str = str.rp(/(\\\([\s\S]+?\\\))/g, protector);
    str = str.rp(/(\\begin\{equation\}[\s\S]*?\\end\{equation\})/g, protector);
    str = str.rp(/(\\begin\{eqnarray\}[\s\S]*?\\end\{eqnarray\})/g, protector);
    str = str.rp(/(\\begin\{equation\*\}[\s\S]*?\\end\{equation\*\})/g, protector);

    // HEADERS
    //
    // We consume leading and trailing whitespace to avoid creating an extra paragraph tag
    // around the header itself.

    // Setext-style H1: Text with ======== right under it
    str = str.rp(/(?:^|\s*\n)(.+?)\n[ \t]*={3,}[ \t]*\n/g, makeHeaderFunc(1));
    
    // Setext-style H2: Text with -------- right under it
    str = str.rp(/(?:^|\s*\n)(.+?)\n[ \t]*-{3,}[ \t]*\n/g, makeHeaderFunc(2));

    // ATX-style headers:
    //
    //  # Foo #
    //  # Foo
    //  (# Bar)
    //
    // If note that '#' in the title are only stripped if they appear at the end, in
    // order to allow headers with # in the title.

    for (var i = 6; i > 0; --i) {
        str = str.rp(new RegExp(/^\s*/.source + '#{' + i + ',' + i +'}(?:[ \t])([^\n]+?)#*[ \t]*\n', 'gm'), 
                 makeHeaderFunc(i));

        // No-number headers
        str = str.rp(new RegExp(/^\s*/.source + '\\(#{' + i + ',' + i +'}\\)(?:[ \t])([^\n]+?)\\(?#*\\)?\\n[ \t]*\n', 'gm'), 
                     '\n</p>\n' + entag('div', '$1', protect('class="nonumberh' + i + '"')) + '\n<p>\n\n');
    }

    // HORIZONTAL RULE: * * *, - - -, _ _ _
    str = str.rp(/\n[ \t]*((\*|-|_)[ \t]*){3,}[ \t]*\n/g, '\n<hr/>\n');

    // PAGE BREAK or HORIZONTAL RULE: +++++
    str = str.rp(/\n[ \t]*\+{5,}[ \t]*\n/g, '\n<hr ' + protect('class="pagebreak"') + '/>\n');

    // ADMONITION: !!! (class) (title)\n body
    str = str.rp(/^!!![ \t]*([^\s"'><&\:]*)\:?(.*)\n([ \t]{3,}.*\s*\n)*/gm, function (match, cssClass, title) {
        // Have to extract the body by splitting match because the regex doesn't capture the body correctly in the multi-line case
        match = match.trim();
        return '\n\n' + entag('div', ((title ? entag('div', title, protect('class="admonition-title"')) + '\n' : '') + match.ss(match.indexOf('\n'))).trim(), protect('class="admonition ' + cssClass.toLowerCase().trim() + '"')) + '\n\n';
    });

    // FANCY QUOTE in a blockquote:
    // > " .... "
    // >    -- Foo

    var FANCY_QUOTE = protect('class="fancyquote"');
    str = str.rp(/\n>[ \t]*"(.*(?:\n>.*)*)"[ \t]*(?:\n>[ \t]*)?(\n>[ \t]{2,}\S.*)?\n/g,
                 function (match, quote, author) {
                     return entag('blockquote', 
                                  entag('span',
                                        quote.rp(/\n>/g, '\n'), 
                                        FANCY_QUOTE) + 
                                  (author ? entag('span',
                                                  author.rp(/\n>/g, '\n'),
                                                  protect('class="author"')) : ''),
                                  FANCY_QUOTE);
                });

    // BLOCKQUOTE: > in front of a series of lines
    // Process iteratively to support nested blockquotes
    var foundBlockquote = false;
    do {
        foundBlockquote = false;
        str = str.rp(/(?:\n>.*){2,}/g, function (match) {
            // Strip the leading '>'
            foundBlockquote = true;
            return entag('blockquote', match.rp(/\n>/g, '\n'));
        });
    } while (foundBlockquote);


    // FOOTNOTES/ENDNOTES: [^symbolic name]. Disallow spaces in footnote names to
    // make parsing unambiguous. Consume leading space before the footnote.
    function endNote(match, symbolicNameA) {
        var symbolicName = symbolicNameA.toLowerCase().trim();

        if (! (symbolicName in endNoteTable)) {
            ++endNoteCount;
            endNoteTable[symbolicName] = endNoteCount;
        }

        return '<sup><a ' + protect('href="#endnote-' + symbolicName + '"') + 
            '>' + endNoteTable[symbolicName] + '</a></sup>';
    }    
    str = str.rp(/[ \t]*\[\^([^\]\n\t ]+)\](?!:)/g, endNote);
    str = str.rp(/(\S)[ \t]*\[\^([^\]\n\t ]+)\]/g, function(match, pre, symbolicNameA) { return pre + endNote(match, symbolicNameA); });


    // CITATIONS: [#symbolicname]
    // The bibliography entry:
    str = str.rp(/\n\[#(\S+)\]:[ \t]+((?:[ \t]*\S[^\n]*\n?)*)/g, function (match, symbolicName, entry) {
        symbolicName = symbolicName.trim();
        return '<div ' + protect('class="bib"') + '>[<a ' + protect('class="target" name="citation-' + symbolicName.toLowerCase() + '"') + 
            '>&nbsp;</a><b>' + symbolicName + '</b>] ' + entry + '</div>';
    });
    
    // A reference:
    // (must process AFTER the definitions, since the syntax is a subset)
    str = str.rp(/\[(#[^\)\(\[\]\.#\s]+(?:\s*,\s*#(?:[^\)\(\[\]\.#\s]+))*)\]/g, function (match, symbolicNameList) {
        // Parse the symbolicNameList
        symbolicNameList = symbolicNameList.split(',');
        var s = '[';
        for (var i = 0; i < symbolicNameList.length; ++i) {
            // Strip spaces and # signs
            var name = symbolicNameList[i].rp(/#| /g, '');
            s += entag('a', name, protect('href="#citation-' + name.toLowerCase() + '"'));
            if (i < symbolicNameList.length - 1) { s += ', '; }
        }
        return s + ']';
    });
    

    // TABLES: line with | over line containing only | and -
    // (process before reference links to avoid ambiguity on the captions)
    str = replaceTables(str, protect);

    // REFERENCE-LINK TABLE: [foo]: http://foo.com
    // (must come before reference images and reference links in processing)
    str = str.rp(/^\[([^\^#].*?)\]:(.*?)$/gm, function (match, symbolicName, url) {
        referenceLinkTable[symbolicName.toLowerCase().trim()] = {link: url.trim(), used: false};
        return '';
    });

    // EMAIL ADDRESS: <foo@bar.baz> or foo@bar.baz if it doesn't look like a URL
    str = str.rp(/(?:<|(?!<)\b)(\S+@(\S+\.)+?\S{2,}?)(?:$|>|(?=<)|(?=\s)(?!>))/g, function (match, addr) {
        if (/http:|ftp:|https:|svn:|:\/\/|\.html|\(|\)|\]/.test(match)) {
            // This is a hyperlink to a url with an @ sign, not an email address
            return match;
        } else {
            return '<a ' + protect('href="mailto:' + addr + '"') + '>' + addr + '</a>';
        }
    });

    // Common code for formatting images
    var formatImage = function (ignore, url, attribs) {
        attribs = attribs || '';
        var img;
        var hash;

        // Detect videos
        if (/\.(mp4|m4v|avi|mpg|mov|webm)$/i.test(url)) {
            // This is video. Any attributes provided will override the defaults given here
            img = '<video ' + protect('class="markdeep" src="' + url + '"' + attribs + ' width="480px" controls="true"') + '></video>';
        } else if (/\.(mp3|mp2|ogg|wav|m4a|aac|flac)$/i.test(url)) {
            // Audio
            img = '<audio ' + protect('class="markdeep" controls ' + attribs + '><source src="' + url + '"') + '></audio>';
        } else if (hash = url.match(/^https:\/\/(?:www\.)?(?:youtube\.com\/\S*?v=|youtu\.be\/)([\w\d-]+)(&.*)?$/i)) {
            // YouTube video
            img = '<iframe ' + protect('class="markdeep" src="https://www.youtube.com/embed/' + hash[1] + '"' + attribs + ' width="480px" height="300px" frameborder="0" allowfullscreen webkitallowfullscreen mozallowfullscreen') + '></iframe>';
        } else if (hash = url.match(/^https:\/\/(?:www\.)?vimeo.com\/\S*?\/([\w\d-]+)$/i)) {
            // Vimeo video
            img = '<iframe ' + protect('class="markdeep" src="https://player.vimeo.com/video/' + hash[1] + '"' + attribs + ' width="480px" height="300px" frameborder="0" allowfullscreen webkitallowfullscreen mozallowfullscreen') + '></iframe>';
        } else {
            // Image (trailing space is needed in case attribs must be quoted by the
            // browser...without the space, the browser will put the closing slash in the
            // quotes.)

            var classList = 'markdeep';
            // Remove classes from attribs
            attribs = attribs.rp(/class *= *(["'])([^'"]+)\1/, function (match, quote, cls) {
                classList += ' ' + cls;
                return '';
            });
            attribs = attribs.rp(/class *= *([^"' ]+)/, function (match, cls) {
                classList += ' ' + cls;
                return '';
            });
            
            img = '<img ' + protect('class="' + classList + '" src="' + url + '"' + attribs) + ' />';
            if (option('autoLinkImages')) {
                img = entag('a', img, protect('href="' + url + '" target="_blank"'));
            }
        }

        return img;
    };

    // Reformat equation links that have brackets: eqn [foo] --> eqn \ref{foo} so that
    // mathjax can process them.
    str = str.rp(/\b(equation|eqn\.|eq\.)\s*\[([^\s\]]+)\]/gi, function (match, eq, label) {
        return eq + ' \\ref{' + label + '}';
    });


    // Reformat figure links that have subfigure labels in parentheses, to avoid them being
    // processed as links
    str = str.rp(/\b(figure|fig\.|table|tbl\.|listing|lst\.)\s*\[([^\s\]]+)\](?=\()/gi, function (match) {
        return match + '<span></span>';
    });


    // Process links before images so that captions can contain links

    // Detect gravizo URLs inside of markdown images and protect them, 
    // which will cause them to be parsed sort-of reasonably. This is
    // a really special case needed to handle the newlines and potential
    // nested parentheses. Use the pattern from http://blog.stevenlevithan.com/archives/regex-recursion
    // (could be extended to multiple nested parens if needed)
    str = str.rp(/\(http:\/\/g.gravizo.com\/(.*g)\?((?:[^\(\)]|\([^\(\)]*\))*)\)/gi, function(match, protocol, url) {
        return "(http://g.gravizo.com/" + protocol + "?" + encodeURIComponent(url) + ")";
    });

    // HYPERLINKS: [text](url attribs)
    str = str.rp(/(^|[^!])\[([^\[\]]+?)\]\(("?)([^<>\s"]+?)\3(\s+[^\)]*?)?\)/g, function (match, pre, text, maybeQuote, url, attribs) {
        attribs = attribs || '';
        return pre + '<a ' + protect('href="' + url + '"' + attribs) + '>' + text + '</a>' + maybeShowLabel(url);
    });

    // EMPTY HYPERLINKS: [](url)
    str = str.rp(/(^|[^!])\[[ \t]*?\]\(("?)([^<>\s"]+?)\2\)/g, function (match, pre, maybeQuote, url) {
        return pre + '<a ' + protect('href="' + url + '"') + '>' + url + '</a>';
    });

    // REFERENCE LINK
    str = str.rp(/(^|[^!])\[([^\[\]]+)\]\[([^\[\]]*)\]/g, function (match, pre, text, symbolicName) {
        // Empty symbolic name is replaced by the label text
        if (! symbolicName.trim()) {
            symbolicName = text;
        }
        
        symbolicName = symbolicName.toLowerCase().trim();
        var t = referenceLinkTable[symbolicName];
        if (! t) {
            console.log("Reference link '" + symbolicName + "' never defined");
            return '?';
        } else {
            t.used = true;
            return pre + '<a ' + protect('href="' + t.link + '"') + '>' + text + '</a>';
        }
    });

    var CAPTION_PROTECT_CHARACTER = '\ue011';
    var protectedCaptionArray = [];
    
    // Temporarily protect image captions (or things that look like
    // them) because the following code is really slow at parsing
    // captions since they have regexps that are complicated to
    // evaluate due to branching.
    //
    // The regexp is really just /.*?\n{0,5}.*/, but that executes substantially more slowly on Chrome.
    str = str.rp(/!\[([^\n\]].*?\n?.*?\n?.*?\n?.*?\n?.*?)\]([\[\(])/g, function (match, caption, bracket) {
        // This is the same as the body of the protect() function, but using the protectedCaptionArray instead
        var i = (protectedCaptionArray.push(caption) - 1).toString(PROTECT_RADIX);
        while (i.length < PROTECT_DIGITS) { i = '0' + i; }
        return '![' + CAPTION_PROTECT_CHARACTER + i + CAPTION_PROTECT_CHARACTER + ']' + bracket;
    });
    
    
    // REFERENCE IMAGE: ![...][ref attribs]
    // Rewrite as a regular image for further processing below.
    str = str.rp(/(!\[.*?\])\[([^<>\[\]\s]+?)([ \t][^\n\[\]]*?)?\]/g, function (match, caption, symbolicName, attribs) {
        symbolicName = symbolicName.toLowerCase().trim();
        var t = referenceLinkTable[symbolicName];
        if (! t) {
            console.log("Reference image '" + symbolicName + "' never defined");
            return '?';
        } else {
            t.used = true;
            var s = caption + '(' + t.link + (t.attribs || '') + ')';
            return s;
        }
    });

    
    // IMAGE GRID: Rewrite rows and grids of images into a grid
    var imageGridAttribs = protect('width="100%"');
    var imageGridRowAttribs = protect('valign="top"');
    // This regex is the pattern for multiple images followed by an optional single image in case the last row is ragged
    // with only one extra
    str = str.rp(/(?:\n(?:[ \t]*!\[.*?\]\(("?)[^<>\s]+?(?:[^\n\)]*?)?\)){2,}[ \t]*)+(?:\n(?:[ \t]*!\[.*?\]\(("?)[^<>\s]+?(?:[^\n\)]*?)?\))[ \t]*)?\n/g, function (match) {
        var table = '';

        // Break into rows:
        match = match.split('\n');

        // Parse each row:
        match.forEach(function(row) {
            row = row.trim();
            if (row) {
                // Parse each image
                table += entag('tr', row.rp(/[ \t]*!\[.*?\]\([^\)\s]+([^\)]*?)?\)/g, function(image, attribs) {
                    //if (! /width|height/i.test(attribs) {
                        // Add a bogus "width" attribute to force the images to be hyperlinked to their
                        // full-resolution versions
                    //}
                    return entag('td', '\n\n'+ image + '\n\n');
                }), imageGridRowAttribs);
            }
        });

        return '\n' + entag('table', table, imageGridAttribs) + '\n';
    });

    // SIMPLE IMAGE: ![](url attribs)
    str = str.rp(/(\s*)!\[\]\(("?)([^"<>\s]+?)\2(\s[^\)]*?)?\)(\s*)/g, function (match, preSpaces, maybeQuote, url, attribs, postSpaces) {
        var img = formatImage(match, url, attribs);

        if (isolated(preSpaces, postSpaces)) {
            // In a block by itself: center
            img = entag('center', img);
        }

        return preSpaces + img + postSpaces;
    });

    // Explicit loop so that the output will be re-processed, preserving spaces between blocks.
    // Note that there is intentionally no global flag on the first regexp since we only want
    // to process the first occurance.
    var loop = true;
    var imageCaptionAbove = option('captionAbove', 'image');
    while (loop) {
        loop = false;

        // CAPTIONED IMAGE: ![caption](url attribs)
        str = str.rp(/(\s*)!\[(.+?)\]\(("?)([^"<>\s]+?)\3(\s[^\)]*?)?\)(\s*)/, function (match, preSpaces, caption, maybeQuote, url, attribs, postSpaces) {
            loop = true;
            var divStyle = '';
            var iso = isolated(preSpaces, postSpaces);

            // Only floating images get their size attributes moved to the whole box
            if (attribs && ! iso) {
                // Move any width *attribute* specification to the box itself
                attribs = attribs.rp(/((?:max-)?width)\s*:\s*[^;'"]*/g, function (attribMatch, attrib) {
                    divStyle = attribMatch + ';';
                    return attrib + ':100%';
                });
                
                // Move any width *style* specification to the box itself
                attribs = attribs.rp(/((?:max-)?width)\s*=\s*('\S+?'|"\S+?")/g, function (attribMatch, attrib, expr) {
                    // Strip the quotes
                    divStyle = attrib + ':' + expr.ss(1, expr.length - 1) + ';';
                    return 'style="' + attrib + ':100%" ';
                });
            }

            var img = formatImage(match, url, attribs);

            if (iso) {
                // In its own block: center
                preSpaces += '<center>';
                postSpaces = '</center>' + postSpaces;
            } else {
                // Embedded: float
                divStyle += 'float:right;margin:4px 0px 0px 25px;'
            }
            var floating = !iso;
            
            caption = entag('center', entag('span', caption + maybeShowLabel(url), protect('class="imagecaption"')));

            // This code used to put floating images in <span> instead of <div>,
            // but it wasn't clear why and this broke centered captions
            return preSpaces + 
                entag('div', (imageCaptionAbove ? caption : '') + img + (! imageCaptionAbove ? caption : ''), protect('class="image" style="' + divStyle + '"')) + 
                postSpaces;
        });
    } // while replacements made

    // Uprotect image captions
    var exposeCaptionRan = false;
    var CAPTION_PROTECT_REGEXP    = RegExp(CAPTION_PROTECT_CHARACTER + '[0-9a-w]{' + PROTECT_DIGITS + ',' + PROTECT_DIGITS + '}' + CAPTION_PROTECT_CHARACTER, 'g');
    /** Given the escaped identifier string from protect(), returns
        the orginal string. */
    function exposeCaption(i) {
        // Strip the escape character and parse, then look up in the
        // dictionary.
        var j = parseInt(i.ss(1, i.length - 1), PROTECT_RADIX);
        exposeCaptionRan = true;
        return protectedCaptionArray[j];
    }
    exposeCaptionRan = true;
    while ((str.indexOf(CAPTION_PROTECT_CHARACTER) + 1) && exposeCaptionRan) {
        exposeCaptionRan = false;
        str = str.rp(CAPTION_PROTECT_REGEXP, exposeCaption);
    }
    
    ////////////////////////////////////////////

    // Process these after links, so that URLs with underscores and tildes are protected.

    // STRONG: Must run before italic, since they use the
    // same symbols. **b** __b__
    str = replaceMatched(str, /\*\*/, 'strong', protect('class="asterisk"'));
    str = replaceMatched(str, /__/, 'strong', protect('class="underscore"'));

    // EM (ITALICS): *i* _i_
    str = replaceMatched(str, /\*/, 'em', protect('class="asterisk"'));
    str = replaceMatched(str, /_/, 'em', protect('class="underscore"'));
    
    // STRIKETHROUGH: ~~text~~
    str = str.rp(/\~\~([^~].*?)\~\~/g, entag('del', '$1'));

    // SMART DOUBLE QUOTES: "a -> localized &ldquo;   z"  -> localized &rdquo;
    // Allow situations such as "foo"==>"bar" and foo:"bar", but not 3' 9"
    if (option('smartQuotes')) {
        str = str.rp(/(^|[ \t->])(")(?=\w)/gm, '$1' + keyword('&ldquo;'));
        str = str.rp(/([A-Za-z\.,:;\?!=<])(")(?=$|\W)/gm, '$1' + keyword('&rdquo;'));
    }
    
    // ARROWS:
    str = str.rp(/(\s|^)<==(\s)/g, '$1\u21D0$2');
    str = str.rp(/(\s|^)->(\s)/g, '$1&rarr;$2');
    // (this requires having removed HTML comments first)
    str = str.rp(/(\s|^)-->(\s)/g, '$1&xrarr;$2');
    str = str.rp(/(\s|^)==>(\s)/g, '$1\u21D2$2');
    str = str.rp(/(\s|^)<-(\s)/g, '$1&larr;$2');
    str = str.rp(/(\s|^)<--(\s)/g, '$1&xlarr;$2');
    str = str.rp(/(\s|^)<==>(\s)/g, '$1\u21D4$2');
    str = str.rp(/(\s|^)<->(\s)/g, '$1\u2194$2');

    // EM DASH: ---
    // (exclude things that look like table delimiters!)
    str = str.rp(/([^-!\:\|])---([^->\:\|])/g, '$1&mdash;$2');

    // other EM DASH: -- (we don't support en dash...it is too short and looks like a minus)
    // (exclude things that look like table delimiters!)
    str = str.rp(/([^-!\:\|])--([^->\:\|])/g, '$1&mdash;$2');

    // NUMBER x NUMBER:
    str = str.rp(/(\d+[ \t]?)x(?=[ \t]?\d+)/g, '$1&times;');

    // MINUS: -4 or 2 - 1
    str = str.rp(/([\s\(\[<\|])-(\d)/g, '$1&minus;$2');
    str = str.rp(/(\d) - (\d)/g, '$1 &minus; $2');

    // EXPONENTS: ^1 ^-1 (no decimal places allowed)
    str = str.rp(/\^([-+]?\d+)\b/g, '<sup>$1</sup>');

    // PAGE BREAK:
    str = str.rp(/(^|\s|\b)\\(pagebreak|newpage)(\b|\s|$)/gi, protect('<div style="page-break-after:always"> </div>\n'))
    
    // SCHEDULE LISTS: date : title followed by indented content
    str = replaceScheduleLists(str, protect);

    // DEFINITION LISTS: Word followed by a colon list
    // Use <dl><dt>term</dt><dd>definition</dd></dl>
    // https://developer.mozilla.org/en-US/docs/Web/HTML/Element/dl
    //
    // Process these before lists so that lists within definition lists
    // work correctly
    str = replaceDefinitionLists(str, protect);

    // LISTS: lines with -, +, *, or number.
    str = replaceLists(str, protect);

    // DEGREE: ##-degree
    str = str.rp(/(\d+?)[ \t-]degree(?:s?)/g, '$1&deg;');

    // PARAGRAPH: Newline, any amount of space, newline...as long as there isn't already
    // a paragraph break there.
    str = str.rp(/(?:<p>)?\n\s*\n+(?!<\/p>)/gi,
                 function(match) { return (/^<p>/i.test(match)) ? match : '\n\n</p><p>\n\n';});

    // Remove empty paragraphs (mostly avoided by the above, but some can still occur)
    str = str.rp(/<p>[\s\n]*<\/p>/gi, '');


    // FOOTNOTES/ENDNOTES
    str = str.rp(/\n\[\^(\S+)\]: ((?:.+?\n?)*)/g, function (match, symbolicName, note) {
        symbolicName = symbolicName.toLowerCase().trim();
        if (symbolicName in endNoteTable) {
            return '\n<div ' + protect('class="endnote"') + '><a ' + 
                protect('class="target" name="endnote-' + symbolicName + '"') + 
                '>&nbsp;</a><sup>' + endNoteTable[symbolicName] + '</sup> ' + note + '</div>';
        } else {
            return "\n";
        }
    });
    

    // SECTION LINKS: XXX section, XXX subsection.
    // Do this by rediscovering the headers and then recursively
    // searching for links to them. Process after other
    // forms of links to avoid ambiguity.
    
    var allHeaders = str.match(/<h([1-6])>(.*?)<\/h\1>/gi);
    if (allHeaders) {
        allHeaders.forEach(function (header) {
            header = removeHTMLTags(header.ss(4, header.length - 5)).trim();
            var link = '<a ' + protect('href="#' + mangle(header) + '"') + '>';

            var sectionExp = '(' + keyword('section') + '|' + keyword('subsection') + '|' + keyword('chapter') + ')';
            var headerExp = '(\\b' + escapeRegExpCharacters(header) + ')';
            
            // Search for links to this section
            str = str.rp(RegExp(headerExp + '\\s+' + sectionExp, 'gi'), link + "$1</a> $2");
            str = str.rp(RegExp(sectionExp + '\\s+' + headerExp, 'gi'), '$1 ' + link + "$2</a>");
        });
    }

    // TABLE, LISTING, and FIGURE LABEL NUMBERING: Figure [symbol]: Table [symbol]: Listing [symbol]: Diagram [symbol]:

    // This data structure maps caption types [by localized name] to a count of how many of
    // that type of object exist.
    var refCounter = {};

    // refTable['type_symbolicName'] = {number: number to link to, used: bool}
    var refTable = {};

    str = str.rp(RegExp(/($|>)\s*/.source + '(' + keyword('figure') + '|' + keyword('table') + '|' + keyword('listing') + '|' + keyword('diagram') + ')' + /\s+\[(.+?)\]:/.source, 'gim'), function (match, prefix, _type, _ref) {
        var type = _type.toLowerCase();
        // Increment the counter
        var count = refCounter[type] = (refCounter[type] | 0) + 1;
        var ref = type + '_' + mangle(_ref.toLowerCase().trim());

        // Store the reference number
        refTable[ref] = {number: count, used: false, source: type + ' [' + _ref + ']'};
        
        return prefix +
               entag('a', '&nbsp;', protect('class="target" name="' + ref + '"')) + entag('b', type[0].toUpperCase() + type.ss(1) + '&nbsp;' + count + ':', protect('style="font-style:normal;"')) +
               maybeShowLabel(_ref);
    });

    // FIGURE, TABLE, DIAGRAM, and LISTING references:
    // (must come after figure/table/listing processing, obviously)
    str = str.rp(RegExp('\\b(fig\\.|tbl\\.|lst\\.|' + keyword('figure') + '|' + keyword('table') + '|' + keyword('listing') + '|' + keyword('diagram') + ')\\s+\\[([^\\s\\]]+)\\]', 'gi'), function (match, _type, _ref) {
        // Fix abbreviations
        var type = _type.toLowerCase();
        switch (type) {
        case 'fig.': type = keyword('figure').toLowerCase(); break;
        case 'tbl.': type = keyword('table').toLowerCase(); break;
        case 'lst.': type = keyword('listing').toLowerCase(); break;
        }

        // Clean up the reference
        var ref = type + '_' + mangle(_ref.toLowerCase().trim());
        var t = refTable[ref];

        if (t) {
            t.used = true;
            return '<a ' + protect('href="#' + ref + '"') + '>' + _type + '&nbsp;' + t.number + maybeShowLabel(_ref) + '</a>';
        } else {
            console.log("Reference to undefined '" + type + " [" + _ref + "]'");
            return _type + ' ?';
        }
    });

    // URL: <http://baz> or http://baz
    // Must be detected after [link]() processing 
    str = str.rp(/(?:<|(?!<)\b)(\w{3,6}:\/\/.+?)(?:$|>|(?=<)|(?=\s|\u00A0)(?!<))/g, function (match, url) {
        var extra = '';
        if (url[url.length - 1] == '.') {
            // Accidentally sucked in a period at the end of a sentence
            url = url.ss(0, url.length - 1);
            extra = '.';
        }
        // svn and perforce URLs are not hyperlinked. All others (http/https/ftp/mailto/tel, etc. are)
        return '<a ' + ((url[0] !== 's' && url[0] !== 'p') ? protect('href="' + url + '" class="url"') : '') + '>' + url + '</a>' + extra;
    });

    if (! elementMode) {
        var TITLE_PATTERN = /^\s*(?:<\/p><p>)?\s*<strong.*?>([^ \t\*].*?[^ \t\*])<\/strong>(?:<\/p>)?[ \t]*\n/.source;
        
        var ALL_SUBTITLES_PATTERN = /([ {4,}\t][ \t]*\S.*\n)*/.source;

        // Detect a bold first line and make it into a title; detect indented lines
        // below it and make them subtitles
        str = str.rp(
            new RegExp(TITLE_PATTERN + ALL_SUBTITLES_PATTERN, 'g'),
            function (match, title) {
                title = title.trim();

                // rp + RegExp won't give us the full list of
                // subtitles, only the last one. So, we have to
                // re-process match.
                var subtitles = match.ss(match.indexOf('\n', match.indexOf('</strong>')));
                subtitles = subtitles ? subtitles.rp(/[ \t]*(\S.*?)\n/g, '<div class="subtitle"> $1 </div>\n') : '';
                
                // Remove all tags from the title when inside the <TITLE> tag, as well
                // as unicode characters that don't render well in tabs and window bars.
                // These regexps look like they are full of spaces but are actually various
                // unicode space characters. http://jkorpela.fi/chars/spaces.html
                var titleTag = removeHTMLTags(title).replace(/[     ]/g, '').replace(/[         　]/g, ' ');
                
                return entag('title', titleTag) + maybeShowLabel(window.location.href, 'center') +
                    '<div class="title"> ' + title + 
                    ' </div>\n' + subtitles + '<div class="afterTitles"></div>\n';
            });
    } // if ! noTitles

    // Remove any bogus leading close-paragraph tag inserted by our extra newlines
    str = str.rp(/^\s*<\/p>/, '');


    // If not in element mode and not an INSERT child, maybe add a TOC
    if (! elementMode) {
        var temp = insertTableOfContents(str, protect, function (text) {return text.rp(PROTECT_REGEXP, expose)});
        str = temp[0];
        var toc = temp[1];
        // SECTION LINKS: Replace sec. [X], section [X], subsection [X]
        str = str.rp(RegExp('\\b(' + keyword('sec') + '\\.|' + keyword('section') + '|' + keyword('subsection') + '|' + keyword('chapter') + ')\\s\\[(.+?)\\]', 'gi'), 
                    function (match, prefix, ref) {
                        var link = toc[ref.toLowerCase().trim()];
                        if (link) {
                            return prefix + '  <a ' + protect('href="#toc' + link + '"') + '>' + link + '</a>';  
                        } else {
                            return prefix + ' ?';
                        }
                    });
    }

    // Expose all protected values. We may need to do this
    // recursively, because pre and code blocks can be nested.
    var maxIterations = 50;

    exposeRan = true;
    while ((str.indexOf(PROTECT_CHARACTER) + 1) && exposeRan && (maxIterations > 0)) {
        exposeRan = false;
        str = str.rp(PROTECT_REGEXP, expose);
        --maxIterations;
    }
    
    if (maxIterations <= 0) { console.log('WARNING: Ran out of iterations while expanding protected substrings'); }

    // Warn about unused references
    Object.keys(referenceLinkTable).forEach(function (key) {
        if (! referenceLinkTable[key].used) {
            console.log("Reference link '[" + key + "]' is defined but never used");
        }
    });

    Object.keys(refTable).forEach(function (key) {
        if (! refTable[key].used) {
            console.log("'" + refTable[key].source + "' is never referenced");
        }
    });

    if (option('linkAPIDefinitions')) {
        // API DEFINITION LINKS
        
        var apiDefined = {};

        // Find link targets for APIs, which look like:
        // '<dt><code...>variablename' followed by (, [, or <
        //
        // If there is syntax highlighting because we're documenting
        // keywords for the language supported by HLJS, then there may
        // be an extra span around the variable name.
        str = str.rp(/<dt><code(\b[^<>\n]*)>(<span class="[a-zA-Z\-_0-9]+">)?([A-Za-z_][A-Za-z_\.0-9:\->]*)(<\/span>)?([\(\[<])/g, function (match, prefix, syntaxHighlight, name, syntaxHighlightEnd, next) {
            var linkName = name + (next === '<' ? '' : next === '(' ? '-fcn' : next === '[' ? '-array' : next);
            apiDefined[linkName] = true;
            // The 'ignore' added to the code tag below is to
            // prevent the link finding code from finding this (since
            // we don't have lookbehinds in JavaScript to recognize
            // the <dt>)
            return '<dt><a name="apiDefinition-' + linkName + '"></a><code ignore ' + prefix + '>' + (syntaxHighlight || '') + name + (syntaxHighlightEnd || '') + next;
        });

        // Hide links that are also inside of a <h#>...</h#>, where we don't want them
        // modified by API links. Assume that these are on a single line. The space in
        // the close tag prevents the next regexp from matching.
        str = str.rp(/<h([1-9])>(.*<code\b[^<>\n]*>.*)<\/code>(.*<\/h\1>)/g, '<h$1>$2</code >$3');

        // Now find potential links, which look like:
        // '<code...>variablename</code>' and may contain () or [] after the variablename
        // They may also have an extra syntax-highlighting span
        str = str.rp(/<code(?! ignore)\b[^<>\n]*>(<span class="[a-zA-Z\-_0-9]+">)?([A-Za-z_][A-Za-z_\.0-9:\->]*)(<\/span>)?(\(\)|\[\])?<\/code>/g, function (match, syntaxHighlight, name, syntaxHighlightEnd, next) {
            var linkName = name + (next ? (next[0] === '(' ? '-fcn' : next[0] === '[' ? '-array' : next[0]) : '');
            return (apiDefined[linkName] === true) ? entag('a', match, 'href="#apiDefinition-' + linkName + '"') : match;
        });
    }
           
    return '<span class="md">' + entag('p', str) + '</span>';
}

 
/** Workaround for IE11 */
function strToArray(s) {
    if (Array.from) {
        return Array.from(s);
    } else {
        var a = [];
        for (var i = 0; i < s.length; ++i) {
            a[i] = s[i];
        }
        return a;
    }
}

/**
   Adds whitespace at the end of each line of str, so that all lines have equal length in
   unicode characters (which is not the same as JavaScript characters when high-index/escape
   characters are present).
*/
function equalizeLineLengths(str) {
    var lineArray = str.split('\n');

    if ((lineArray.length > 0) && (lineArray[lineArray.length - 1] === '')) {
        // Remove the empty last line generated by split on a trailing newline
        lineArray.pop();
    }
        
    var longest = 0;
    lineArray.forEach(function(line) {
        longest = max(longest, strToArray(line).length);
    });

    // Worst case spaces needed for equalizing lengths
    // http://stackoverflow.com/questions/1877475/repeat-character-n-times
    var spaces = Array(longest + 1).join(' ');

    var result = '';
    lineArray.forEach(function(line) {
        // Append the needed number of spaces onto each line, and
        // reconstruct the output with newlines
        result += line + spaces.ss(strToArray(line).length) + '\n';
    });

    return result;
}

/** Finds the longest common whitespace prefix of all non-empty lines
    and then removes it */
function removeLeadingSpace(str) {
    var lineArray = str.split('\n');

    var minimum = Infinity;
    lineArray.forEach(function (line) {
        if (line.trim() !== '') {
            // This is a non-empty line
            var spaceArray = line.match(/^([ \t]*)/);
            if (spaceArray) {
                minimum = min(minimum, spaceArray[0].length);
            }
        }
    });

    if (minimum === 0) {
        // No leading space
        return str;
    }

    var result = '';
    lineArray.forEach(function(line) {
        // Strip the common spaces
        result += line.ss(minimum) + '\n';
    });

    return result;
}

/** Returns true if this character is a "letter" under the ASCII definition */
function isASCIILetter(c) {
    var code = c.charCodeAt(0);
    return ((code >= 65) && (code <= 90)) || ((code >= 97) && (code <= 122));
}

/** Converts diagramString, which is a Markdeep diagram without the surrounding asterisks, to
    SVG (HTML). Lines may have ragged lengths.

    alignmentHint is the float alignment desired for the SVG tag,
    which can be 'floatleft', 'floatright', or ''
 */
function diagramToSVG(diagramString, alignmentHint) {
    // Clean up diagramString if line endings are ragged
    diagramString = equalizeLineLengths(diagramString);

    // Temporarily replace 'o' that is surrounded by other text
    // with another character to avoid processing it as a point 
    // decoration. This will be replaced in the final svg and is
    // faster than checking each neighborhood each time.
    var HIDE_O = '\ue004';
    diagramString = diagramString.rp(/([a-zA-Z]{2})o/g, '$1' + HIDE_O);
    diagramString = diagramString.rp(/o([a-zA-Z]{2})/g, HIDE_O + '$1');
    diagramString = diagramString.rp(/([a-zA-Z\ue004])o([a-zA-Z\ue004])/g, '$1' + HIDE_O + '$2');

    /** Pixels per character */
    var SCALE   = 8;

    /** Multiply Y coordinates by this when generating the final SVG
        result to account for the aspect ratio of text files. This
        MUST be 2 */
    var ASPECT = 2;

    var DIAGONAL_ANGLE = Math.atan(1.0 / ASPECT) * 180 / Math.PI;

    var EPSILON = 1e-6;

    // The order of the following is based on rotation angles
    // and is used for ArrowSet.toSVG
    var ARROW_HEAD_CHARACTERS            = '>v<^';
    var POINT_CHARACTERS                 = 'o*◌○◍●';
    var JUMP_CHARACTERS                  = '()';
    var UNDIRECTED_VERTEX_CHARACTERS     = "+";
    var VERTEX_CHARACTERS                = UNDIRECTED_VERTEX_CHARACTERS + ".'";

    // GRAY[i] is the Unicode block character for (i+1)/4 level gray
    var GRAY_CHARACTERS = '\u2591\u2592\u2593\u2588';

    // TRI[i] is a right-triangle rotated by 90*i
    var TRI_CHARACTERS  = '\u25E2\u25E3\u25E4\u25E5';

    var DECORATION_CHARACTERS            = ARROW_HEAD_CHARACTERS + POINT_CHARACTERS + JUMP_CHARACTERS + GRAY_CHARACTERS + TRI_CHARACTERS;

    function isUndirectedVertex(c) { return UNDIRECTED_VERTEX_CHARACTERS.indexOf(c) + 1; }
    function isVertex(c)           { return VERTEX_CHARACTERS.indexOf(c) !== -1; }
    function isTopVertex(c)        { return isUndirectedVertex(c) || (c === '.'); }
    function isBottomVertex(c)     { return isUndirectedVertex(c) || (c === "'"); }
    function isVertexOrLeftDecoration(c){ return isVertex(c) || (c === '<') || isPoint(c); }
    function isVertexOrRightDecoration(c){return isVertex(c) || (c === '>') || isPoint(c); }
    function isArrowHead(c)        { return ARROW_HEAD_CHARACTERS.indexOf(c) + 1; }
    function isGray(c)             { return GRAY_CHARACTERS.indexOf(c) + 1; }
    function isTri(c)              { return TRI_CHARACTERS.indexOf(c) + 1; }

    // "D" = Diagonal slash (/), "B" = diagonal Backslash (\)
    // Characters that may appear anywhere on a solid line
    function isSolidHLine(c)       { return (c === '-') || isUndirectedVertex(c) || isJump(c); }
    function isSolidVLineOrJumpOrPoint(c) { return isSolidVLine(c) || isJump(c) || isPoint(c); }
    function isSolidVLine(c)       { return (c === '|') || isUndirectedVertex(c); }
    function isSolidDLine(c)       { return (c === '/') || isUndirectedVertex(c) }
    function isSolidBLine(c)       { return (c === '\\') || isUndirectedVertex(c); }
    function isJump(c)             { return JUMP_CHARACTERS.indexOf(c) + 1; }
    function isPoint(c)            { return POINT_CHARACTERS.indexOf(c) + 1; }
    function isDecoration(c)       { return DECORATION_CHARACTERS.indexOf(c) + 1; }
    function isEmpty(c)            { return c === ' '; }
   
    ///////////////////////////////////////////////////////////////////////////////
    // Math library

    /** Invoke as new Vec2(v) to clone or new Vec2(x, y) to create from coordinates.
        Can also invoke without new for brevity. */
    function Vec2(x, y) {
        // Detect when being run without new
        if (! (this instanceof Vec2)) { return new Vec2(x, y); }

        if (y === undefined) {
            if (x === undefined) { x = y = 0; } 
            else if (x instanceof Vec2) { y = x.y; x = x.x; }
            else { console.error("Vec2 requires one Vec2 or (x, y) as an argument"); }
        }
        this.x = x;
        this.y = y;
        Object.seal(this);
    }

    /** Returns an SVG representation */
    Vec2.prototype.toString = Vec2.prototype.toSVG = 
        function () { return '' + (this.x * SCALE) + ',' + (this.y * SCALE * ASPECT) + ' '; };

    /** Converts a "rectangular" string defined by newlines into 2D
        array of characters. Grids are immutable. */
    function makeGrid(str) {
        /** Returns ' ' for out of bounds values */
        var grid = function(x, y) {
            if (y === undefined) {
                if (x instanceof Vec2) { y = x.y; x = x.x; }
                else { console.error('grid requires either a Vec2 or (x, y)'); }
            }
            
            return ((x >= 0) && (x < grid.width) && (y >= 0) && (y < grid.height)) ?
                str[y * (grid.width + 1) + x] : ' ';
        };

        // Elements are true when consumed
        grid._used   = [];

        grid.height  = str.split('\n').length;
        if (str[str.length - 1] === '\n') { --grid.height; }

        // Convert the string to an array to better handle greater-than 16-bit unicode
        // characters, which JavaScript does not process correctly with indices. Do this after
        // the above string processing.
        str = strToArray(str);
        grid.width = str.indexOf('\n');

        /** Mark this location. Takes a Vec2 or (x, y) */
        grid.setUsed = function (x, y) {
            if (y === undefined) {
                if (x instanceof Vec2) { y = x.y; x = x.x; }
                else { console.error('grid requires either a Vec2 or (x, y)'); }
            }
            if ((x >= 0) && (x < grid.width) && (y >= 0) && (y < grid.height)) {
                // Match the source string indexing
                grid._used[y * (grid.width + 1) + x] = true;
            }
        };
        
        grid.isUsed = function (x, y) {
            if (y === undefined) {
                if (x instanceof Vec2) { y = x.y; x = x.x; }
                else { console.error('grid requires either a Vec2 or (x, y)'); }
            }
            return (this._used[y * (this.width + 1) + x] === true);
        };
        
        /** Returns true if there is a solid vertical line passing through (x, y) */
        grid.isSolidVLineAt = function (x, y) {
            if (y === undefined) { y = x.x; x = x.x; }
            
            var up = grid(x, y - 1);
            var c  = grid(x, y);
            var dn = grid(x, y + 1);
            
            var uprt = grid(x + 1, y - 1);
            var uplt = grid(x - 1, y - 1);
            
            if (isSolidVLine(c)) {
                // Looks like a vertical line...does it continue?
                return (isTopVertex(up)    || (up === '^') || isSolidVLine(up) || isJump(up) ||
                        isBottomVertex(dn) || (dn === 'v') || isSolidVLine(dn) || isJump(dn) ||
                        isPoint(up) || isPoint(dn) || (grid(x, y - 1) === '_') || (uplt === '_') ||
                        (uprt === '_') ||
                        
                        // Special case of 1-high vertical on two curved corners 
                        ((isTopVertex(uplt) || isTopVertex(uprt)) &&
                         (isBottomVertex(grid(x - 1, y + 1)) || isBottomVertex(grid(x + 1, y + 1)))));
                
            } else if (isTopVertex(c) || (c === '^')) {
                // May be the top of a vertical line
                return isSolidVLine(dn) || (isJump(dn) && (c !== '.'));
            } else if (isBottomVertex(c) || (c === 'v')) {
                return isSolidVLine(up) || (isJump(up) && (c !== "'"));
            } else if (isPoint(c)) {
                return isSolidVLine(up) || isSolidVLine(dn);
            } 
            
            return false;
        };
    
    
        /** Returns true if there is a solid middle (---) horizontal line
            passing through (x, y). Ignores underscores. */
        grid.isSolidHLineAt = function (x, y) {
            if (y === undefined) { y = x.x; x = x.x; }
            
            var ltlt = grid(x - 2, y);
            var lt   = grid(x - 1, y);
            var c    = grid(x + 0, y);
            var rt   = grid(x + 1, y);
            var rtrt = grid(x + 2, y);
            
            if (isSolidHLine(c) || (isSolidHLine(lt) && isJump(c))) {
                // Looks like a horizontal line...does it continue? We need three in a row.
                if (isSolidHLine(lt)) {
                    return isSolidHLine(rt) || isVertexOrRightDecoration(rt) || 
                        isSolidHLine(ltlt) || isVertexOrLeftDecoration(ltlt);
                } else if (isVertexOrLeftDecoration(lt)) {
                    return isSolidHLine(rt);
                } else {
                    return isSolidHLine(rt) && (isSolidHLine(rtrt) || isVertexOrRightDecoration(rtrt));
                }

            } else if (c === '<') {
                return isSolidHLine(rt) && isSolidHLine(rtrt)
                
            } else if (c === '>') {
                return isSolidHLine(lt) && isSolidHLine(ltlt);
                
            } else if (isVertex(c)) {
                return ((isSolidHLine(lt) && isSolidHLine(ltlt)) || 
                        (isSolidHLine(rt) && isSolidHLine(rtrt)));
            }
            
            return false;
        };
        
        
        /** Returns true if there is a solid backslash line passing through (x, y) */
        grid.isSolidBLineAt = function (x, y) {
            if (y === undefined) { y = x.x; x = x.x; }
            var c = grid(x, y);
            var lt = grid(x - 1, y - 1);
            var rt = grid(x + 1, y + 1);
            
            if (c === '\\') {
                // Looks like a diagonal line...does it continue? We need two in a row.
                return (isSolidBLine(rt) || isBottomVertex(rt) || isPoint(rt) || (rt === 'v') ||
                        isSolidBLine(lt) || isTopVertex(lt) || isPoint(lt) || (lt === '^') ||
                        (grid(x, y - 1) === '/') || (grid(x, y + 1) === '/') || (rt === '_') || (lt === '_')); 
            } else if (c === '.') {
                return (rt === '\\');
            } else if (c === "'") {
                return (lt === '\\');
            } else if (c === '^') {
                return rt === '\\';
            } else if (c === 'v') {
                return lt === '\\';
            } else if (isVertex(c) || isPoint(c) || (c === '|')) {
                return isSolidBLine(lt) || isSolidBLine(rt);
            }
        };
        

        /** Returns true if there is a solid diagonal line passing through (x, y) */
        grid.isSolidDLineAt = function (x, y) {
            if (y === undefined) { y = x.x; x = x.x; }
            
            var c = grid(x, y);
            var lt = grid(x - 1, y + 1);
            var rt = grid(x + 1, y - 1);
            
            if (c === '/' && ((grid(x, y - 1) === '\\') || (grid(x, y + 1) === '\\'))) {
                // Special case of tiny hexagon corner
                return true;
            } else if (isSolidDLine(c)) {
                // Looks like a diagonal line...does it continue? We need two in a row.
                return (isSolidDLine(rt) || isTopVertex(rt) || isPoint(rt) || (rt === '^') || (rt === '_') ||
                        isSolidDLine(lt) || isBottomVertex(lt) || isPoint(lt) || (lt === 'v') || (lt === '_')); 
            } else if (c === '.') {
                return (lt === '/');
            } else if (c === "'") {
                return (rt === '/');
            } else if (c === '^') {
                return lt === '/';
            } else if (c === 'v') {
                return rt === '/';
            } else if (isVertex(c) || isPoint(c) || (c === '|')) {
                return isSolidDLine(lt) || isSolidDLine(rt);
            }
            return false;
        };
        
        grid.toString = function () { return str; };
        
        return Object.freeze(grid);
    }
    
    
    /** A 1D curve. If C is specified, the result is a bezier with
        that as the tangent control point */
    function Path(A, B, C, D, dashed) {
        if (! ((A instanceof Vec2) && (B instanceof Vec2))) {
            console.error('Path constructor requires at least two Vec2s');
        }
        this.A = A;
        this.B = B;
        if (C) {
            this.C = C;
            if (D) {
                this.D = D;
            } else {
                this.D = C;
            }
        }

        this.dashed = dashed || false;

        Object.freeze(this);
    }

    var _ = Path.prototype;
    _.isVertical = function () {
        return this.B.x === this.A.x;
    };

    _.isHorizontal = function () {
        return this.B.y === this.A.y;
    };

    /** Diagonal lines look like: / See also backDiagonal */
    _.isDiagonal = function () {
        var dx = this.B.x - this.A.x;
        var dy = this.B.y - this.A.y;
        return (abs(dy + dx) < EPSILON);
    };

    _.isBackDiagonal = function () {
        var dx = this.B.x - this.A.x;
        var dy = this.B.y - this.A.y;
        return (abs(dy - dx) < EPSILON);
    };

    _.isCurved = function () {
        return this.C !== undefined;
    };

    /** Does this path have any end at (x, y) */
    _.endsAt = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return ((this.A.x === x) && (this.A.y === y)) ||
            ((this.B.x === x) && (this.B.y === y));
    };

    /** Does this path have an up end at (x, y) */
    _.upEndsAt = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return this.isVertical() && (this.A.x === x) && (min(this.A.y, this.B.y) === y);
    };

    /** Does this path have an up end at (x, y) */
    _.diagonalUpEndsAt = function (x, y) {
        if (! this.isDiagonal()) { return false; }
        if (y === undefined) { y = x.y; x = x.x; }
        if (this.A.y < this.B.y) {
            return (this.A.x === x) && (this.A.y === y);
        } else {
            return (this.B.x === x) && (this.B.y === y);
        }
    };

    /** Does this path have a down end at (x, y) */
    _.diagonalDownEndsAt = function (x, y) {
        if (! this.isDiagonal()) { return false; }
        if (y === undefined) { y = x.y; x = x.x; }
        if (this.B.y < this.A.y) {
            return (this.A.x === x) && (this.A.y === y);
        } else {
            return (this.B.x === x) && (this.B.y === y);
        }
    };

    /** Does this path have an up end at (x, y) */
    _.backDiagonalUpEndsAt = function (x, y) {
        if (! this.isBackDiagonal()) { return false; }
        if (y === undefined) { y = x.y; x = x.x; }
        if (this.A.y < this.B.y) {
            return (this.A.x === x) && (this.A.y === y);
        } else {
            return (this.B.x === x) && (this.B.y === y);
        }
    };

    /** Does this path have a down end at (x, y) */
    _.backDiagonalDownEndsAt = function (x, y) {
        if (! this.isBackDiagonal()) { return false; }
        if (y === undefined) { y = x.y; x = x.x; }
        if (this.B.y < this.A.y) {
            return (this.A.x === x) && (this.A.y === y);
        } else {
            return (this.B.x === x) && (this.B.y === y);
        }
    };

    /** Does this path have a down end at (x, y) */
    _.downEndsAt = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return this.isVertical() && (this.A.x === x) && (max(this.A.y, this.B.y) === y);
    };

    /** Does this path have a left end at (x, y) */
    _.leftEndsAt = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return this.isHorizontal() && (this.A.y === y) && (min(this.A.x, this.B.x) === x);
    };

    /** Does this path have a right end at (x, y) */
    _.rightEndsAt = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return this.isHorizontal() && (this.A.y === y) && (max(this.A.x, this.B.x) === x);
    };

    _.verticalPassesThrough = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return this.isVertical() && 
            (this.A.x === x) && 
            (min(this.A.y, this.B.y) <= y) &&
            (max(this.A.y, this.B.y) >= y);
    }

    _.horizontalPassesThrough = function (x, y) {
        if (y === undefined) { y = x.y; x = x.x; }
        return this.isHorizontal() && 
            (this.A.y === y) && 
            (min(this.A.x, this.B.x) <= x) &&
            (max(this.A.x, this.B.x) >= x);
    }
    
    /** Returns a string suitable for inclusion in an SVG tag */
    _.toSVG = function () {
        var svg = '<path d="M ' + this.A;

        if (this.isCurved()) {
            svg += 'C ' + this.C + this.D + this.B;
        } else {
            svg += 'L ' + this.B;
        }
        svg += '" style="fill:none;"';
        if (this.dashed) {
            svg += ' stroke-dasharray="3,6"';
        }
        svg += '/>';
        return svg;
    };


    /** A group of 1D curves. This was designed so that all of the
        methods can later be implemented in O(1) time, but it
        currently uses O(n) implementations for source code
        simplicity. */
    function PathSet() {
        this._pathArray = [];
    }

    var PS = PathSet.prototype;
    PS.insert = function (path) {
        this._pathArray.push(path);
    };

    /** Returns a new method that returns true if method(x, y) 
        returns true on any element of _pathAray */
    function makeFilterAny(method) {
        return function(x, y) {
            for (var i = 0; i < this._pathArray.length; ++i) {
                if (method.call(this._pathArray[i], x, y)) { return true; }
            }
            // Fall through: return undefined == false
        }
    }

    // True if an up line ends at these coordinates. Recall that the
    // variable _ is bound to the Path prototype still.
    PS.upEndsAt                = makeFilterAny(_.upEndsAt);
    PS.diagonalUpEndsAt        = makeFilterAny(_.diagonalUpEndsAt);
    PS.backDiagonalUpEndsAt    = makeFilterAny(_.backDiagonalUpEndsAt);
    PS.diagonalDownEndsAt      = makeFilterAny(_.diagonalDownEndsAt);
    PS.backDiagonalDownEndsAt  = makeFilterAny(_.backDiagonalDownEndsAt);
    PS.downEndsAt              = makeFilterAny(_.downEndsAt);
    PS.leftEndsAt              = makeFilterAny(_.leftEndsAt);
    PS.rightEndsAt             = makeFilterAny(_.rightEndsAt);
    PS.endsAt                  = makeFilterAny(_.endsAt);
    PS.verticalPassesThrough   = makeFilterAny(_.verticalPassesThrough);
    PS.horizontalPassesThrough = makeFilterAny(_.horizontalPassesThrough);

    /** Returns an SVG string */
    PS.toSVG = function () {
        var svg = '';
        for (var i = 0; i < this._pathArray.length; ++i) {
            svg += this._pathArray[i].toSVG() + '\n';
        }
        return svg;
    };


    function DecorationSet() {
        this._decorationArray = [];
    }

    var DS = DecorationSet.prototype;

    /** insert(x, y, type, <angle>)  
        insert(vec, type, <angle>)

        angle is the angle in degrees to rotate the result */
    DS.insert = function(x, y, type, angle) {
        if (type === undefined) { type = y; y = x.y; x = x.x; }

        if (! isDecoration(type)) {
            console.error('Illegal decoration character: ' + type); 
        }
        var d = {C: Vec2(x, y), type: type, angle:angle || 0};

        // Put arrows at the front and points at the back so that
        // arrows always draw under points

        if (isPoint(type)) {
            this._decorationArray.push(d);
        } else {
            this._decorationArray.unshift(d);
        }
    };


    DS.toSVG = function () {
        var svg = '';
        for (var i = 0; i < this._decorationArray.length; ++i) {
            var decoration = this._decorationArray[i];
            var C = decoration.C;
            
            if (isJump(decoration.type)) {
                // Slide jumps
                var dx = (decoration.type === ')') ? +0.75 : -0.75;
                var up  = Vec2(C.x, C.y - 0.5);
                var dn  = Vec2(C.x, C.y + 0.5);
                var cup = Vec2(C.x + dx, C.y - 0.5);
                var cdn = Vec2(C.x + dx, C.y + 0.5);

                svg += '<path d="M ' + dn + ' C ' + cdn + cup + up + '" style="fill:none;"/>';

            } else if (isPoint(decoration.type)) {
                var cls = {'*':'closed', 'o':'open', '◌':'dotted', '○':'open', '◍':'shaded', '●':'closed'}[decoration.type];
                svg += '<circle cx="' + (C.x * SCALE) + '" cy="' + (C.y * SCALE * ASPECT) +
                       '" r="' + (SCALE - STROKE_WIDTH) + '" class="' + cls + 'dot"/>';
            } else if (isGray(decoration.type)) {
                
                var shade = Math.round((3 - GRAY_CHARACTERS.indexOf(decoration.type)) * 63.75);
                svg += '<rect x="' + ((C.x - 0.5) * SCALE) + '" y="' + ((C.y - 0.5) * SCALE * ASPECT) + '" width="' + SCALE + '" height="' + (SCALE * ASPECT) + '" stroke="none" fill="rgb(' + shade + ',' + shade + ',' + shade +')"/>';

            } else if (isTri(decoration.type)) {
                // 30-60-90 triangle
                var index = TRI_CHARACTERS.indexOf(decoration.type);
                var xs  = 0.5 - (index & 1);
                var ys  = 0.5 - (index >> 1);
                xs *= sign(ys);
                var tip = Vec2(C.x + xs, C.y - ys);
                var up  = Vec2(C.x + xs, C.y + ys);
                var dn  = Vec2(C.x - xs, C.y + ys);
                svg += '<polygon points="' + tip + up + dn + '" style="stroke:none"/>\n';
            } else { // Arrow head
                var tip = Vec2(C.x + 1, C.y);
                var up =  Vec2(C.x - 0.5, C.y - 0.35);
                var dn =  Vec2(C.x - 0.5, C.y + 0.35);
                svg += '<polygon points="' + tip + up + dn + 
                    '"  style="stroke:none" transform="rotate(' + decoration.angle + ',' + C + ')"/>\n';
            }
        }
        return svg;
    };

    ////////////////////////////////////////////////////////////////////////////

    function findPaths(grid, pathSet) {
        // Does the line from A to B contain at least one c?
        function lineContains(A, B, c) {
            var dx = sign(B.x - A.x);
            var dy = sign(B.y - A.y);
            var x, y;

            for (x = A.x, y = A.y; (x !== B.x) || (y !== B.y); x += dx, y += dy) {
                if (grid(x, y) === c) { return true; }
            }

            // Last point
            return (grid(x, y) === c);
        }

        // Find all solid vertical lines. Iterate horizontally
        // so that we never hit the same line twice
        for (var x = 0; x < grid.width; ++x) {
            for (var y = 0; y < grid.height; ++y) {
                if (grid.isSolidVLineAt(x, y)) {
                    // This character begins a vertical line...now, find the end
                    var A = Vec2(x, y);
                    do  { grid.setUsed(x, y); ++y; } while (grid.isSolidVLineAt(x, y));
                    var B = Vec2(x, y - 1);
                    
                    var up = grid(A);
                    var upup = grid(A.x, A.y - 1);

                    if (! isVertex(up) && ((upup === '-') || (upup === '_') ||
                                           (upup === '┳') ||
                                           (grid(A.x - 1, A.y - 1) === '_') ||
                                           (grid(A.x + 1, A.y - 1) === '_') || 
                                           isBottomVertex(upup)) || isJump(upup)) {
                        // Stretch up to almost reach the line above (if there is a decoration,
                        // it will finish the gap)
                        A.y -= 0.5;
                    }

                    var dn = grid(B);
                    var dndn = grid(B.x, B.y + 1);
                    if (! isVertex(dn) && ((dndn === '-') || (dndn === '┻') || isTopVertex(dndn)) || isJump(dndn) ||
                        (grid(B.x - 1, B.y) === '_') || (grid(B.x + 1, B.y) === '_')) {
                        // Stretch down to almost reach the line below
                        B.y += 0.5;
                    }

                    // Don't insert degenerate lines
                    if ((A.x !== B.x) || (A.y !== B.y)) {
                        pathSet.insert(new Path(A, B));
                    }

                    // Continue the search from the end value y+1
                } 

                // Some very special patterns for the short lines needed on
                // circuit diagrams. Only invoke these if not also on a curve
                //      _  _    
                //    -'    '-   -'
                else if ((grid(x, y) === "'") &&
                    (((grid(x - 1, y) === '-') && (grid(x + 1, y - 1) === '_') &&
                     ! isSolidVLineOrJumpOrPoint(grid(x - 1, y - 1))) ||
                     ((grid(x - 1, y - 1) === '_') && (grid(x + 1, y) === '-') &&
                     ! isSolidVLineOrJumpOrPoint(grid(x + 1, y - 1))))) {
                    pathSet.insert(new Path(Vec2(x, y - 0.5), Vec2(x, y)));
                }

                //    _.-  -._  
                else if ((grid(x, y) === '.') &&
                         (((grid(x - 1, y) === '_') && (grid(x + 1, y) === '-') && 
                           ! isSolidVLineOrJumpOrPoint(grid(x + 1, y + 1))) ||
                          ((grid(x - 1, y) === '-') && (grid(x + 1, y) === '_') &&
                           ! isSolidVLineOrJumpOrPoint(grid(x - 1, y + 1))))) {
                    pathSet.insert(new Path(Vec2(x, y), Vec2(x, y + 0.5)));
                }

                // For drawing resistors: -.╱
                else if ((grid(x, y) === '.') &&
                         (grid(x - 1, y) === '-') &&
                         (grid(x + 1, y) === '╱')) {
                    pathSet.insert(new Path(Vec2(x, y), Vec2(x + 0.5, y + 0.5)));
                }
                
                // For drawing resistors: ╱'-
                else if ((grid(x, y) === "'") &&
                         (grid(x + 1, y) === '-') &&
                         (grid(x - 1, y) === '╱')) {
                    pathSet.insert(new Path(Vec2(x, y), Vec2(x - 0.5, y - 0.5)));
                }

            } // y
        } // x
        
        // Find all solid horizontal lines 
        for (var y = 0; y < grid.height; ++y) {
            for (var x = 0; x < grid.width; ++x) {
                if (grid.isSolidHLineAt(x, y)) {
                    // Begins a line...find the end
                    var A = Vec2(x, y);
                    do { grid.setUsed(x, y); ++x; } while (grid.isSolidHLineAt(x, y));
                    var B = Vec2(x - 1, y);

                    // Detect adjacent box-drawing characters and lengthen the edge
                    if (grid(B.x + 1, B.y) === '┫') { B.x += 0.5; }
                    if (grid(A.x - 1, A.y) === '┣') { A.x -= 0.5; }

                    // Detect curves and shorten the edge
                    if ( ! isVertex(grid(A.x - 1, A.y)) && 
                         ((isTopVertex(grid(A)) && isSolidVLineOrJumpOrPoint(grid(A.x - 1, A.y + 1))) ||
                          (isBottomVertex(grid(A)) && isSolidVLineOrJumpOrPoint(grid(A.x - 1, A.y - 1))))) {
                        ++A.x;
                    }

                    if ( ! isVertex(grid(B.x + 1, B.y)) && 
                         ((isTopVertex(grid(B)) && isSolidVLineOrJumpOrPoint(grid(B.x + 1, B.y + 1))) ||
                          (isBottomVertex(grid(B)) && isSolidVLineOrJumpOrPoint(grid(B.x + 1, B.y - 1))))) {
                        --B.x;
                    }

                    // Only insert non-degenerate lines
                    if ((A.x !== B.x) || (A.y !== B.y)) {
                        pathSet.insert(new Path(A, B));
                    }
                    
                    // Continue the search from the end x+1
                }
            }
        } // y

        // Find all solid left-to-right downward diagonal lines (BACK DIAGONAL)
        for (var i = -grid.height; i < grid.width; ++i) {
            for (var x = i, y = 0; y < grid.height; ++y, ++x) {
                if (grid.isSolidBLineAt(x, y)) {
                    // Begins a line...find the end
                    var A = Vec2(x, y);
                    do { ++x; ++y; } while (grid.isSolidBLineAt(x, y));
                    var B = Vec2(x - 1, y - 1);

                    // Ensure that the entire line wasn't just vertices
                    if (lineContains(A, B, '\\')) {
                        for (var j = A.x; j <= B.x; ++j) {
                            grid.setUsed(j, A.y + (j - A.x)); 
                        }

                        var top = grid(A);
                        var up = grid(A.x, A.y - 1);
                        var uplt = grid(A.x - 1, A.y - 1);
                        if ((up === '/') || (uplt === '_') || (up === '_') || 
                            (! isVertex(top)  && 
                             (isSolidHLine(uplt) || isSolidVLine(uplt)))) {
                            // Continue half a cell more to connect for:
                            //  ___   ___
                            //  \        \    /      ----     |
                            //   \        \   \        ^      |^
                            A.x -= 0.5; A.y -= 0.5;
                        } else if (isPoint(uplt)) {
                            // Continue 1/4 cell more to connect for:
                            //
                            //  o
                            //   ^
                            //    \
                            A.x -= 0.25; A.y -= 0.25;
                        }
                        
                        var bottom = grid(B);
                        var dnrt = grid(B.x + 1, B.y + 1);
                        if ((grid(B.x, B.y + 1) === '/') || (grid(B.x + 1, B.y) === '_') || 
                            (grid(B.x - 1, B.y) === '_') || 
                            (! isVertex(grid(B)) &&
                             (isSolidHLine(dnrt) || isSolidVLine(dnrt)))) {
                            // Continue half a cell more to connect for:
                            //                       \      \ |
                            //  \       \     \       v      v|
                            //   \__   __\    /      ----     |
                            
                            B.x += 0.5; B.y += 0.5;
                        } else if (isPoint(dnrt)) {
                            // Continue 1/4 cell more to connect for:
                            //
                            //    \
                            //     v
                            //      o
                            
                            B.x += 0.25; B.y += 0.25;
                        }
                        
                        pathSet.insert(new Path(A, B));
                        // Continue the search from the end x+1,y+1
                    } // lineContains
                }
            }
        } // i


        // Find all solid left-to-right upward diagonal lines (DIAGONAL)
        for (var i = -grid.height; i < grid.width; ++i) {
            for (var x = i, y = grid.height - 1; y >= 0; --y, ++x) {
                if (grid.isSolidDLineAt(x, y)) {
                    // Begins a line...find the end
                    var A = Vec2(x, y);
                    do { ++x; --y; } while (grid.isSolidDLineAt(x, y));
                    var B = Vec2(x - 1, y + 1);

                    if (lineContains(A, B, '/')) {
                        // This is definitely a line. Commit the characters on it
                        for (var j = A.x; j <= B.x; ++j) {
                            grid.setUsed(j, A.y - (j - A.x)); 
                        }

                        var up = grid(B.x, B.y - 1);
                        var uprt = grid(B.x + 1, B.y - 1);
                        var bottom = grid(B);
                        if ((up === '\\') || (up === '_') || (uprt === '_') || 
                            (! isVertex(grid(B)) &&
                             (isSolidHLine(uprt) || isSolidVLine(uprt)))) {
                            
                            // Continue half a cell more to connect at:
                            //     __   __  ---     |
                            //    /      /   ^     ^|
                            //   /      /   /     / |
                            
                            B.x += 0.5; B.y -= 0.5;
                        } else if (isPoint(uprt)) {
                            
                            // Continue 1/4 cell more to connect at:
                            //
                            //       o
                            //      ^
                            //     /
                            
                            B.x += 0.25; B.y -= 0.25;
                        }
                        
                        var dnlt = grid(A.x - 1, A.y + 1);
                        var top = grid(A);
                        if ((grid(A.x, A.y + 1) === '\\') || (grid(A.x - 1, A.y) === '_') || (grid(A.x + 1, A.y) === '_') ||
                            (! isVertex(grid(A)) &&
                             (isSolidHLine(dnlt) || isSolidVLine(dnlt)))) {

                            // Continue half a cell more to connect at:
                            //               /     \ |
                            //    /  /      v       v|
                            // __/  /__   ----       | 
                            
                            A.x -= 0.5; A.y += 0.5;
                        } else if (isPoint(dnlt)) {
                            
                            // Continue 1/4 cell more to connect at:
                            //
                            //       /
                            //      v
                            //     o
                            
                            A.x -= 0.25; A.y += 0.25;
                        }
                        pathSet.insert(new Path(A, B));

                        // Continue the search from the end x+1,y-1
                    } // lineContains
                }
            }
        } // y
        
        
        // Now look for curved corners. The syntax constraints require
        // that these can always be identified by looking at three
        // horizontally-adjacent characters.
        for (var y = 0; y < grid.height; ++y) {
            for (var x = 0; x < grid.width; ++x) {
                var c = grid(x, y);

                // Note that because of undirected vertices, the
                // following cases are not exclusive
                if (isTopVertex(c)) {
                    // -.
                    //   |
                    if (isSolidHLine(grid(x - 1, y)) && isSolidVLine(grid(x + 1, y + 1))) {
                        grid.setUsed(x - 1, y); grid.setUsed(x, y); grid.setUsed(x + 1, y + 1);
                        pathSet.insert(new Path(Vec2(x - 1, y), Vec2(x + 1, y + 1), 
                                                Vec2(x + 1.1, y), Vec2(x + 1, y + 1)));
                    }

                    //  .-
                    // |
                    if (isSolidHLine(grid(x + 1, y)) && isSolidVLine(grid(x - 1, y + 1))) {
                        grid.setUsed(x - 1, y + 1); grid.setUsed(x, y); grid.setUsed(x + 1, y);
                        pathSet.insert(new Path(Vec2(x + 1, y), Vec2(x - 1, y + 1), 
                                                Vec2(x - 1.1, y), Vec2(x - 1, y + 1)));
                    }
                }
                
                // Special case patterns:
                //   .  .   .  .    
                //  (  o     )  o
                //   '  .   '  '
                if (((c === ')') || isPoint(c)) && (grid(x - 1, y - 1) === '.') && (grid(x - 1, y + 1) === "\'")) {
                    grid.setUsed(x, y); grid.setUsed(x - 1, y - 1); grid.setUsed(x - 1, y + 1);
                    pathSet.insert(new Path(Vec2(x - 2, y - 1), Vec2(x - 2, y + 1), 
                                            Vec2(x + 0.6, y - 1), Vec2(x + 0.6, y + 1)));
                }

                if (((c === '(') || isPoint(c)) && (grid(x + 1, y - 1) === '.') && (grid(x + 1, y + 1) === "\'")) {
                    grid.setUsed(x, y); grid.setUsed(x + 1, y - 1); grid.setUsed(x + 1, y + 1);
                    pathSet.insert(new Path(Vec2(x + 2, y - 1), Vec2(x + 2, y + 1), 
                                            Vec2(x - 0.6, y - 1), Vec2(x - 0.6, y + 1)));
                }

                if (isBottomVertex(c)) {
                    //   |
                    // -' 
                    if (isSolidHLine(grid(x - 1, y)) && isSolidVLine(grid(x + 1, y - 1))) {
                        grid.setUsed(x - 1, y); grid.setUsed(x, y); grid.setUsed(x + 1, y - 1);
                        pathSet.insert(new Path(Vec2(x - 1, y), Vec2(x + 1, y - 1), 
                                                Vec2(x + 1.1, y), Vec2(x + 1, y - 1)));
                    }

                    // | 
                    //  '-
                    if (isSolidHLine(grid(x + 1, y)) && isSolidVLine(grid(x - 1, y - 1))) {
                        grid.setUsed(x - 1, y - 1); grid.setUsed(x, y); grid.setUsed(x + 1, y);
                        pathSet.insert(new Path(Vec2(x + 1, y), Vec2(x - 1, y - 1),
                                                Vec2(x - 1.1, y), Vec2(x - 1, y - 1)));
                    }
                }
               
            } // for x
        } // for y

        // Find low horizontal lines marked with underscores. These
        // are so simple compared to the other cases that we process
        // them directly here without a helper function. Process these
        // from top to bottom and left to right so that we can read
        // them in a single sweep.
        // 
        // Exclude the special case of double underscores going right
        // into an ASCII character, which could be a source code
        // identifier such as __FILE__ embedded in the diagram.
        for (var y = 0; y < grid.height; ++y) {
            for (var x = 0; x < grid.width - 2; ++x) {
                var lt = grid(x - 1, y);

                if ((grid(x, y) === '_') && (grid(x + 1, y) === '_') && 
                    (! isASCIILetter(grid(x + 2, y)) || (lt === '_')) && 
                    (! isASCIILetter(lt) || (grid(x + 2, y) === '_'))) {

                    var ltlt = grid(x - 2, y);
                    var A = Vec2(x - 0.5, y + 0.5);

                    if ((lt === '|') || (grid(x - 1, y + 1) === '|') ||
                        (lt === '.') || (grid(x - 1, y + 1) === "'")) {
                        // Extend to meet adjacent vertical
                        A.x -= 0.5;

                        // Very special case of overrunning into the side of a curve,
                        // needed for logic gate diagrams
                        if ((lt === '.') && 
                            ((ltlt === '-') ||
                             (ltlt === '.')) &&
                            (grid(x - 2, y + 1) === '(')) {
                            A.x -= 0.5;
                        }
                    } else if (lt === '/') {
                        A.x -= 1.0;
                    }

                    // Detect overrun of a tight double curve
                    if ((lt === '(') && (ltlt === '(') &&
                        (grid(x, y + 1) === "'") && (grid(x, y - 1) === '.')) {
                        A.x += 0.5;
                    }
                    lt = ltlt = undefined;

                    do { grid.setUsed(x, y); ++x; } while (grid(x, y) === '_');

                    var B = Vec2(x - 0.5, y + 0.5);
                    var c = grid(x, y);
                    var rt = grid(x + 1, y);
                    var dn = grid(x, y + 1);

                    if ((c === '|') || (dn === '|') || (c === '.') || (dn === "'")) {
                        // Extend to meet adjacent vertical
                        B.x += 0.5;

                        // Very special case of overrunning into the side of a curve,
                        // needed for logic gate diagrams
                        if ((c === '.') && 
                            ((rt === '-') || (rt === '.')) &&
                            (grid(x + 1, y + 1) === ')')) {
                            B.x += 0.5;
                        }
                    } else if ((c === '\\')) {
                        B.x += 1.0;
                    }

                    // Detect overrun of a tight double curve
                    if ((c === ')') && (rt === ')') && (grid(x - 1, y + 1) === "'") && (grid(x - 1, y - 1) === '.')) {
                        B.x += -0.5;
                    }

                    pathSet.insert(new Path(A, B));
                }
            } // for x
        } // for y
    } // findPaths


    function findDecorations(grid, pathSet, decorationSet) {
        function isEmptyOrVertex(c) { return (c === ' ') || /[^a-zA-Z0-9]|[ov]/.test(c); }
        function isLetter(c) { var x = c.toUpperCase().charCodeAt(0); return (x > 64) && (x < 91); }
                    
        /** Is the point in the center of these values on a line? Allow points that are vertically
            adjacent but not horizontally--they wouldn't fit anyway, and might be text. */
        function onLine(up, dn, lt, rt) {
            return ((isEmptyOrVertex(dn) || isPoint(dn)) &&
                    (isEmptyOrVertex(up) || isPoint(up)) &&
                    isEmptyOrVertex(rt) &&
                    isEmptyOrVertex(lt));
        }

        for (var x = 0; x < grid.width; ++x) {
            for (var j = 0; j < grid.height; ++j) {
                var c = grid(x, j);
                var y = j;

                if (isJump(c)) {

                    // Ensure that this is really a jump and not a stray character
                    if (pathSet.downEndsAt(x, y - 0.5) &&
                        pathSet.upEndsAt(x, y + 0.5)) {
                        decorationSet.insert(x, y, c);
                        grid.setUsed(x, y);
                    }

                } else if (isPoint(c)) {
                    var up = grid(x, y - 1);
                    var dn = grid(x, y + 1);
                    var lt = grid(x - 1, y);
                    var rt = grid(x + 1, y);
                    var llt = grid(x - 2, y);
                    var rrt = grid(x + 2, y);

                    if (pathSet.rightEndsAt(x - 1, y) ||   // Must be at the end of a line...
                        pathSet.leftEndsAt(x + 1, y) ||    // or completely isolated NSEW
                        pathSet.downEndsAt(x, y - 1) ||
                        pathSet.upEndsAt(x, y + 1) ||

                        pathSet.upEndsAt(x, y) ||    // For points on vertical lines 
                        pathSet.downEndsAt(x, y) ||  // that are surrounded by other characters
                        
                        onLine(up, dn, lt, rt)) {

                        decorationSet.insert(x, y, c);
                        grid.setUsed(x, y);
                    }
                } else if (isGray(c)) {
                    decorationSet.insert(x, y, c);
                    grid.setUsed(x, y);
                } else if (isTri(c)) {
                    decorationSet.insert(x, y, c);
                    grid.setUsed(x, y);
                } else { // Arrow heads

                    // If we find one, ensure that it is really an
                    // arrow head and not a stray character by looking
                    // for a connecting line.
                    var dx = 0;
                    if ((c === '>') && (pathSet.rightEndsAt(x, y) || 
                                        pathSet.horizontalPassesThrough(x, y))) {
                        if (isPoint(grid(x + 1, y))) {
                            // Back up if connecting to a point so as to not
                            // overlap it
                            dx = -0.5;
                        }
                        decorationSet.insert(x + dx, y, '>', 0);
                        grid.setUsed(x, y);
                    } else if ((c === '<') && (pathSet.leftEndsAt(x, y) ||
                                               pathSet.horizontalPassesThrough(x, y))) {
                        if (isPoint(grid(x - 1, y))) {
                            // Back up if connecting to a point so as to not
                            // overlap it
                            dx = 0.5;
                        }
                        decorationSet.insert(x + dx, y, '>', 180); 
                        grid.setUsed(x, y);
                    } else if (c === '^') {
                        // Because of the aspect ratio, we need to look
                        // in two slots for the end of the previous line
                        if (pathSet.upEndsAt(x, y - 0.5)) {
                            decorationSet.insert(x, y - 0.5, '>', 270); 
                            grid.setUsed(x, y);
                        } else if (pathSet.upEndsAt(x, y)) {
                            decorationSet.insert(x, y, '>', 270);
                            grid.setUsed(x, y);
                        } else if (pathSet.diagonalUpEndsAt(x + 0.5, y - 0.5)) {
                            decorationSet.insert(x + 0.5, y - 0.5, '>', 270 + DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.diagonalUpEndsAt(x + 0.25, y - 0.25)) {
                            decorationSet.insert(x + 0.25, y - 0.25, '>', 270 + DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.diagonalUpEndsAt(x, y)) {
                            decorationSet.insert(x, y, '>', 270 + DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.backDiagonalUpEndsAt(x, y)) {
                            decorationSet.insert(x, y, c, 270 - DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.backDiagonalUpEndsAt(x - 0.5, y - 0.5)) {
                            decorationSet.insert(x - 0.5, y - 0.5, c, 270 - DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.backDiagonalUpEndsAt(x - 0.25, y - 0.25)) {
                            decorationSet.insert(x - 0.25, y - 0.25, c, 270 - DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.verticalPassesThrough(x, y)) {
                            // Only try this if all others failed
                            decorationSet.insert(x, y - 0.5, '>', 270); 
                            grid.setUsed(x, y);
                        }
                    } else if (c === 'v') {
                        if (pathSet.downEndsAt(x, y + 0.5)) {
                            decorationSet.insert(x, y + 0.5, '>', 90); 
                            grid.setUsed(x, y);
                        } else if (pathSet.downEndsAt(x, y)) {
                            decorationSet.insert(x, y, '>', 90);
                            grid.setUsed(x, y);
                        } else if (pathSet.diagonalDownEndsAt(x, y)) {
                            decorationSet.insert(x, y, '>', 90 + DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.diagonalDownEndsAt(x - 0.5, y + 0.5)) {
                            decorationSet.insert(x - 0.5, y + 0.5, '>', 90 + DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.diagonalDownEndsAt(x - 0.25, y + 0.25)) {
                            decorationSet.insert(x - 0.25, y + 0.25, '>', 90 + DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.backDiagonalDownEndsAt(x, y)) {
                            decorationSet.insert(x, y, '>', 90 - DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.backDiagonalDownEndsAt(x + 0.5, y + 0.5)) {
                            decorationSet.insert(x + 0.5, y + 0.5, '>', 90 - DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.backDiagonalDownEndsAt(x + 0.25, y + 0.25)) {
                            decorationSet.insert(x + 0.25, y + 0.25, '>', 90 - DIAGONAL_ANGLE);
                            grid.setUsed(x, y);
                        } else if (pathSet.verticalPassesThrough(x, y)) {
                            // Only try this if all others failed
                            decorationSet.insert(x, y + 0.5, '>', 90); 
                            grid.setUsed(x, y);
                        }
                    } // arrow heads
                } // decoration type
            } // y
        } // x
    } // findArrowHeads

    // Cases where we want to redraw at graphical unicode character
    // to adjust its weight or shape for a conventional application
    // in constructing a diagram.
    function findReplacementCharacters(grid, pathSet) {
        for (var x = 0; x < grid.width; ++x) {
            for (var y = 0; y < grid.height; ++y) {
                if (grid.isUsed(x, y)) continue;
                var c = grid(x, y);
                switch (c) {
                case '╱':
                    pathSet.insert(new Path(Vec2(x - 0.5, y + 0.5), Vec2(x + 0.5, y - 0.5)));
                    grid.setUsed(x, y);
                    break;
                case '╲':
                    pathSet.insert(new Path(Vec2(x - 0.5, y - 0.5), Vec2(x + 0.5, y + 0.5)));
                    grid.setUsed(x, y);
                    break;
                }
            }
        }
    } // findReplacementCharacters

    var grid = makeGrid(diagramString);

    var pathSet = new PathSet();
    var decorationSet = new DecorationSet();

    findPaths(grid, pathSet);
    findReplacementCharacters(grid, pathSet);
    findDecorations(grid, pathSet, decorationSet);

    var svg = '<svg class="diagram" xmlns="http://www.w3.org/2000/svg" version="1.1" height="' + 
        ((grid.height + 1) * SCALE * ASPECT) + '" width="' + ((grid.width + 1) * SCALE) + '"';

    if (alignmentHint === 'floatleft') {
        svg += ' style="float:left;margin:15px 30px 15px 0;"';
    } else if (alignmentHint === 'floatright') {
        svg += ' style="float:right;margin:15px 0 15px 30px;"';
    } else if (alignmentHint === 'center') {
        svg += ' style="margin:0 auto 0 auto;"';
    }

    svg += '><g transform="translate(' + Vec2(1, 1) + ')">\n';

    if (DEBUG_SHOW_GRID) {
        svg += '<g style="opacity:0.1">\n';
        for (var x = 0; x < grid.width; ++x) {
            for (var y = 0; y < grid.height; ++y) {
                svg += '<rect x="' + ((x - 0.5) * SCALE + 1) + '" + y="' + ((y - 0.5) * SCALE * ASPECT + 2) + '" width="' + (SCALE - 2) + '" height="' + (SCALE * ASPECT - 2) + '" style="fill:';
                if (grid.isUsed(x, y)) {
                    svg += 'red;';
                } else if (grid(x, y) === ' ') {
                    svg += 'gray;opacity:0.05';
                } else {
                    svg += 'blue;';
                }
                svg += '"/>\n';
            }
        }
        svg += '</g>\n';
    }
    
    svg += pathSet.toSVG();
    svg += decorationSet.toSVG();

    // Convert any remaining characters
    if (! DEBUG_HIDE_PASSTHROUGH) {
        svg += '<g transform="translate(0,0)">';
        for (var y = 0; y < grid.height; ++y) {
            for (var x = 0; x < grid.width; ++x) {
                var c = grid(x, y);
                if (/[\u2B22\u2B21]/.test(c)) {
                    // Enlarge hexagons so that they fill a grid
                    svg += '<text text-anchor="middle" x="' + (x * SCALE) + '" y="' + (4 + y * SCALE * ASPECT) + '" style="font-size:20.5px">' + escapeHTMLEntities(c) +  '</text>';
                } else if ((c !== ' ') && ! grid.isUsed(x, y)) {
                    svg += '<text text-anchor="middle" x="' + (x * SCALE) + '" y="' + (4 + y * SCALE * ASPECT) + '">' + escapeHTMLEntities(c) +  '</text>';
                } // if
            } // y
        } // x
        svg += '</g>';
    }

    if (DEBUG_SHOW_SOURCE) {
        // Offset the characters a little for easier viewing
        svg += '<g transform="translate(2,2)">\n';
        for (var x = 0; x < grid.width; ++x) {
            for (var y = 0; y < grid.height; ++y) {
                var c = grid(x, y);
                if (c !== ' ') {
                    svg += '<text text-anchor="middle" x="' + (x * SCALE) + '" y="' + (4 + y * SCALE * ASPECT) + '" style="fill:#F00;font-family:Menlo,monospace;font-size:12px;text-align:center">' + escapeHTMLEntities(c) +  '</text>';
                } // if
            } // y
        } // x
        svg += '</g>';
    } // if

    svg += '</g></svg>';

    svg = svg.rp(new RegExp(HIDE_O, 'g'), 'o');


    return svg;
}


////////////////////////// Processing for INSERT HERE
//
// Insert command processing modifies the entire document and potentially
// delays further processing, so it is handled specially and runs the main
// markdeep processing as a callback
//
// node: the node being processed for markdeep. This is document.body
// in markdeep mode, but may be another node in html or script mode.
//
// processMarkdeepCallback: function to run when insert is complete
// to evaluate markdeep 
function processInsertCommands(nodeArray, sourceArray, insertDoneCallback) {
    var myURLParse = /([^?]+)(?:\?id=(inc\d+)&p=([^&]+))?/.exec(location.href);

    var myBase = removeFilename(myURLParse[1]);
    var myID = myURLParse[2];
    var parentBase = removeFilename(myURLParse[3] && decodeURIComponent(myURLParse[3]));
    var childFrameStyle = 'display:none';
    var includeCounter = 0;
    var IAmAChild = myID; // !== undefined
    var IAmAParent = false;
    var numIncludeChildrenLeft = 0;
    
    // Helper function for use by children
    function sendContentsToMyParent() {
        var body = document.body.innerHTML;

        // Fix relative URLs within the body
        var baseref;
        if (document.baseURI !== undefined) {
            baseref = document.baseURI.rp(/\/[^/]+$/, '/');
        } else {
            // IE11
            // Return location from BASE tag.
            //   https://developer.mozilla.org/en-US/docs/Web/HTML/Element/base
            var base = document.getElementsByTagName('base');
            baseref = (base.length > 0) ? base[0].href : document.URL;
        }

        var serverref;
        if (/^file:\/\//.test(baseref)) {
            serverref = 'file://';
        } else {
            serverref = baseref.match(/[^:/]{3,6}:\/\/[^/]*\//)[0];
        }

        // Cases where URLs appear:
        //
        // ![](...)
        // [](...)
        // [link]: ...
        // <img src="...">
        // <script src="...">
        // <a href="...">
        // <link href="...">
        //
        // A url is relative if it does not begin with '^[a-z]{3,6}://|^#'

        // Protect code fences
        // TODO

        function makeAbsoluteURL(url) {
            return (/^[a-z]{3,6}:\/\//.test(url)) ?
                url :
                (url[0] === '/') ?
                // Make relative to server
                serverref + url.ss(1) :
                // Make relative to source document
                baseref + url;
        }

        // Unquoted images and links
        body = body.rp(/\]\([ \t]*([^#")][^ "\)]+)([ \t\)])/g, function (match, url, suffix) {
            return '](' + makeAbsoluteURL(url) + suffix;
        });
        
        // Quoted images and links
        body = body.rp(/\]\([ \t]*"([^#"][^"]+)"([ \t\)])/g, function (match, url, suffix) {
            return ']("' + makeAbsoluteURL(url) + '"' + suffix;
        });

        // Raw HTML
        body = body.rp(/(src|href)=(["'])([^#>][^"'\n>]+)\2/g, function (match, type, quot, url) {
            return type + '=' + quot + makeAbsoluteURL(url) + quot;
        });

        // Reference links
        body = body.rp(/(\n\[[^\]>\n \t]:[ \t]*)([^# \t][^ \t]+)"/g, function (match, prefix, url) {
            return prefix + makeAbsoluteURL(url);
        });

        // Unprotect code fences
        // TODO
        
        // console.log(location.pathname + " sent message to parent");
        // Send the document contents after the childFrame replaced itself
        // (not the source variable captured when this function was defined!)
        parent.postMessage([myID, '=', body].join(''), '*');
    }

    // Strip the filename from the url, if there is one (and it is a string)
    function removeFilename(url) {
        return url && url.ss(0, url.lastIndexOf('/') + 1);
    }

    // Called when this entire document is ready for either markdeep
    // processing or sending to its parent for markdeep processing.
    //
    // IAmAChild: Truish if this document is a child
    //
    // sourceArray: If known, source is the code for the nodes. If it was modified, it is not provided
    function documentReady(IAmAChild, nodeArray, sourceArray) {
        if (IAmAChild) {
            // I'm a child and not waiting for my own children, so trigger the send now. My parent will
            // do the processing.
            
            // console.log("Leaf node " + location.pathname + " sending to parent");
            sendContentsToMyParent();
        } else {
            // No includes. Run markdeep processing after the rest of this file parses
            
            // console.log("non-parent, non-child Parent scheduling markdeepProcessor");
            setTimeout(function () { insertDoneCallback(nodeArray, sourceArray) }, 1);
        }
    }
     
     function messageCallback(event) {
         // Parse the message. Ensure that it is for the Markdeep/include.js system.
         var childID = false;
         var childBody = event.data.substring && event.data.replace(/^(inc\d+)=/, function (match, a) {
             childID = a;
             return '';
         });
         
         if (childID) {
             // This message event was for the Markdeep/include.js system
             
             //console.log(location.href + ' received a message from child ' + childID);

             // Replace the corresponding node's contents
             var childFrame = document.getElementById(childID);
             childFrame.outerHTML = '\n' + childBody + '\n';

             --numIncludeChildrenLeft;

             //console.log(window.location.pathname, 'numIncludeChildrenLeft = ' + numIncludeChildrenLeft);
             
             if (numIncludeChildrenLeft <= 0) {
                 // This was the last child
                 documentReady(IAmAChild, nodeArray);
             }
         }
     };

     var isFirefox = navigator.userAgent.indexOf('Firefox') !== -1 && navigator.userAgent.indexOf('Seamonkey') === -1;
    
     // Find all insert or embed statements in all nodes and replace them
     for (var i = 0; i < sourceArray.length; ++i) {
         sourceArray[i] = sourceArray[i].rp(/(?:^|\s)\((insert|embed)[ \t]+(\S+\.\S*)[ \t]+(height=[a-zA-Z0-9.]+[ \t]+)?here\)\s/g, function(match, type, src, params) {
             var childID = 'inc' + (++includeCounter);
             var isHTML = src.toLowerCase().rp(/\?.*$/,'').endsWith('.html');
             if (type === 'embed' || ! isHTML) {
                 // This is not embedding another Markdeep file. Instead it is embedding
                 // some other kind of document.
                 var tag = 'iframe', url='src';
                 var style = params ? ' style="' + params.rp(/=/g, ':') + '"' : '';
                 
                 if (isFirefox && ! isHTML) {
                     // Firefox doesn't handle embedding other non-html documents in iframes
                     // correctly (it tries to download them!), so we switch to an object
                     // tag--which seems to work identically to the embed tag on this browser.                     
                     tag = 'object'; url = 'data';

                     // Firefox can be confused on a server (but not
                     // locally) by misconfigured MIME types and show
                     // nothing.  But if we know that we're on a
                     // server, we can go ahead and make an
                     // XMLHttpRequest() for the underlying document
                     // directly. Replace the insert in this case.
                     if (location.protocol.startsWith('http')) {
                         var req = new XMLHttpRequest();
                         (function (childID, style) {
                             req.addEventListener("load", function () {
                                 document.getElementById(childID).outerHTML =
                                     entag('iframe', '', 'class="textinsert" srcdoc="<pre>' + this.responseText.replace(/"/g, '&quot;') + '</pre>"' + style);
                             });
                             req.overrideMimeType("text/plain; charset=x-user-defined");
                             req.open("GET", src); 
                             req.send();
                         })(childID, style);
                     }
                 }

                 return entag(tag, '', 'class="textinsert" id="' + childID + '" ' + url + '="' + src + '"' + style);
             }
             
             if (numIncludeChildrenLeft === 0) {
                 // This is the first child observed. Prepare to receive messages from the
                 // embedded children.
                 IAmAParent = true;
                 addEventListener("message", messageCallback);
             }
             
             ++numIncludeChildrenLeft;
             //console.log(window.location.pathname, 'numIncludeChildrenLeft = ' + numIncludeChildrenLeft);
             
             // Replace this tag with a frame that loads the document.  Once loaded, it will
             // send a message with its contents for use as a replacement.
             return '<iframe src="' + src + '?id=' + childID + '&p=' + encodeURIComponent(myBase) + 
                 '" id="' + childID + '" style="' + childFrameStyle + '" content="text/html;charset=UTF-8"></iframe>';
         });
     }

     // console.log('after insert: ' + source);

     // Process all nodes
     if (IAmAParent) {
         // I'm waiting on children, so don't run the full processor
         // yet, but do substitute the iframe code so that it can
         // launch. I may be a child as well...this will be determined
         // when numIncludeChildren hits zero.

         for (var i = 0; i < sourceArray.length; ++i) {
             nodeArray[i].innerHTML = sourceArray[i];
         }
     } else {
         // The source was not modified
         documentReady(IAmAChild, nodeArray, sourceArray);
     }
} // function processInsertCommands()

 
/* xcode.min.js modified */
var HIGHLIGHT_STYLESHEET =
        "<style>.hljs{display:block;overflow-x:auto;padding:0.5em;background:#fff;color:#000;-webkit-text-size-adjust:none}"+
        ".hljs-comment{color:#006a00}" +
        ".hljs-keyword{color:#02E}" +
        ".hljs-literal,.nginx .hljs-title{color:#aa0d91}" + 
        ".method,.hljs-list .hljs-title,.hljs-tag .hljs-title,.setting .hljs-value,.hljs-winutils,.tex .hljs-command,.http .hljs-title,.hljs-request,.hljs-status,.hljs-name{color:#008}" + 
        ".hljs-envvar,.tex .hljs-special{color:#660}" + 
        ".hljs-string{color:#c41a16}" +
        ".hljs-tag .hljs-value,.hljs-cdata,.hljs-filter .hljs-argument,.hljs-attr_selector,.apache .hljs-cbracket,.hljs-date,.hljs-regexp{color:#080}" + 
        ".hljs-sub .hljs-identifier,.hljs-pi,.hljs-tag,.hljs-tag .hljs-keyword,.hljs-decorator,.ini .hljs-title,.hljs-shebang,.hljs-prompt,.hljs-hexcolor,.hljs-rule .hljs-value,.hljs-symbol,.hljs-symbol .hljs-string,.hljs-number,.css .hljs-function,.hljs-function .hljs-title,.coffeescript .hljs-attribute{color:#A0C}" +
        ".hljs-function .hljs-title{font-weight:bold;color:#000}" + 
        ".hljs-class .hljs-title,.smalltalk .hljs-class,.hljs-type,.hljs-typename,.hljs-tag .hljs-attribute,.hljs-doctype,.hljs-class .hljs-id,.hljs-built_in,.setting,.hljs-params,.clojure .hljs-attribute{color:#5c2699}" +
        ".hljs-variable{color:#3f6e74}" +
        ".css .hljs-tag,.hljs-rule .hljs-property,.hljs-pseudo,.hljs-subst{color:#000}" + 
        ".css .hljs-class,.css .hljs-id{color:#9b703f}" +
        ".hljs-value .hljs-important{color:#ff7700;font-weight:bold}" +
        ".hljs-rule .hljs-keyword{color:#c5af75}" +
        ".hljs-annotation,.apache .hljs-sqbracket,.nginx .hljs-built_in{color:#9b859d}" +
        ".hljs-preprocessor,.hljs-preprocessor *,.hljs-pragma{color:#643820}" +
        ".tex .hljs-formula{background-color:#eee;font-style:italic}" +
        ".diff .hljs-header,.hljs-chunk{color:#808080;font-weight:bold}" +
        ".diff .hljs-change{background-color:#bccff9}" +
        ".hljs-addition{background-color:#baeeba}" +
        ".hljs-deletion{background-color:#ffc8bd}" +
        ".hljs-comment .hljs-doctag{font-weight:bold}" +
        ".method .hljs-id{color:#000}</style>";

function isMarkdeepScriptName(str) { return str.search(/markdeep\S*?\.js$/i) !== -1; }
function toArray(list) { return Array.prototype.slice.call(list); }

// Intentionally uninitialized global variable used to detect
// recursive invocations
if (! window.alreadyProcessedMarkdeep) {
    window.alreadyProcessedMarkdeep = true;

    // Detect the noformat argument to the URL
    var noformat = (window.location.href.search(/\?.*noformat.*/i) !== -1);

    // Export relevant methods
    window.markdeep = Object.freeze({ 
        format:               markdeepToHTML,
        formatDiagram:        diagramToSVG,
        stylesheet:           function() {
            return STYLESHEET + sectionNumberingStylesheet() + HIGHLIGHT_STYLESHEET;
        }
    });

    // Not needed: jax: ["input/TeX", "output/SVG"], 
    var MATHJAX_CONFIG ='<script type="text/x-mathjax-config">MathJax.Hub.Config({ TeX: { equationNumbers: {autoNumber: "AMS"} } });</script>' +
        '<span style="display:none">' +
        // Custom definitions (NC == \newcommand)
        '$$NC{\\n}{\\hat{n}}NC{\\thetai}{\\theta_\\mathrm{i}}NC{\\thetao}{\\theta_\\mathrm{o}}NC{\\d}[1]{\\mathrm{d}#1}NC{\\w}{\\hat{\\omega}}NC{\\wi}{\\w_\\mathrm{i}}NC{\\wo}{\\w_\\mathrm{o}}NC{\\wh}{\\w_\\mathrm{h}}NC{\\Li}{L_\\mathrm{i}}NC{\\Lo}{L_\\mathrm{o}}NC{\\Le}{L_\\mathrm{e}}NC{\\Lr}{L_\\mathrm{r}}NC{\\Lt}{L_\\mathrm{t}}NC{\\O}{\\mathrm{O}}NC{\\degrees}{{^{\\large\\circ}}}NC{\\T}{\\mathsf{T}}NC{\\mathset}[1]{\\mathbb{#1}}NC{\\Real}{\\mathset{R}}NC{\\Integer}{\\mathset{Z}}NC{\\Boolean}{\\mathset{B}}NC{\\Complex}{\\mathset{C}}NC{\\un}[1]{\\,\\mathrm{#1}}$$\n'.rp(/NC/g, '\\newcommand') +
        '</span>\n';

    // The following option forces better rendering on some browsers, but also makes it impossible to copy-paste text with
    // inline equations:
    //
    // 'config=TeX-MML-AM_SVG'
    var MATHJAX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.6/MathJax.js?config=TeX-AMS-MML_HTMLorMML';

    var loadMathJax = function() {
        // Dynamically load mathjax
        var script = document.createElement("script");
        script.type = "text/javascript";
        script.src = MATHJAX_URL;
        document.getElementsByTagName("head")[0].appendChild(script);
    }

    var needsMathJax= function(html) {
        // Need MathJax if $$ ... $$, \( ... \), or \begin{
        return option('detectMath') &&
            ((html.search(/(?:\$\$[\s\S]+\$\$)|(?:\\begin{)/m) !== -1) || 
             (html.search(/\\\(.*\\\)/) !== -1));
    }

    var mode = option('mode');
    switch (mode) {
    case 'script':
        // Nothing to do
        return;

    case 'html':
    case 'doxygen':
        // Process explicit diagram tags by themselves
        toArray(document.getElementsByClassName('diagram')).concat(toArray(document.getElementsByTagName('diagram'))).forEach(
            function (element) {
                var src = unescapeHTMLEntities(element.innerHTML);
                // Remove the first and last string (which probably
                // had the pre or diagram tag as part of them) if they are 
                // empty except for whitespace.
                src = src.rp(/(:?^[ \t]*\n)|(:?\n[ \t]*)$/g, '');

                if (mode === 'doxygen') {
                    // Undo Doxygen's &ndash and &mdash, which are impossible to 
                    // detect once the browser has parsed the document
                    src = src.rp(new RegExp('\u2013', 'g'), '--');
                    src = src.rp(new RegExp('\u2014', 'g'), '---');
                    
                    // Undo Doxygen's links within the diagram because they throw off spacing
                    src = src.rp(/<a class="el" .*>(.*)<\/a>/g, '$1');
                }
                element.outerHTML = '<center class="md">' + diagramToSVG(removeLeadingSpace(src), '') + '</center>';
            });

        // Collect all nodes that will receive Markdeep processing
        var markdeepNodeArray = toArray(document.getElementsByClassName('markdeep')).concat(toArray(document.getElementsByTagName('markdeep')));

        // Extract the source code of markeep nodes
        var sourceArray = markdeepNodeArray.map(function (node) {
            return removeLeadingSpace(unescapeHTMLEntities(node.innerHTML));
        });

        // Process insert commands and then trigger markdeep processing
        processInsertCommands(markdeepNodeArray, sourceArray, function (nodeArray, sourceArray) {
            // Update sourceArray if needed because the source code was mutated
            // by insert processing
            sourceArray = sourceArray || nodeArray.map(function (node) {
                return removeLeadingSpace(unescapeHTMLEntities(node.innerHTML));
            });
            
            // Process all nodes, replacing them as we progress
            var anyNeedsMathJax = false;
            for (var i = 0; i < markdeepNodeArray.length; ++i) {
                var oldNode = markdeepNodeArray[i];
                var newNode = document.createElement('div');
                var source = removeLeadingSpace(unescapeHTMLEntities(oldNode.innerHTML));
                var html = markdeepToHTML(source, true);
                anyNeedsMathJax = anyNeedsMathJax || needsMathJax(html);
                newNode.innerHTML = html;
                oldNode.parentNode.replaceChild(newNode, oldNode);
            }

            if (anyNeedsMathJax) { loadMathJax(); }

            // Include our stylesheet even if there are no MARKDEEP tags, but do not include the BODY_STYLESHEET.
            document.head.innerHTML = window.markdeep.stylesheet() + document.head.innerHTML + (anyNeedsMathJax ? MATHJAX_CONFIG : '');

            // Remove fallback nodes
            var fallbackNodes = document.getElementsByClassName('fallback');
            for (var i = 0; i < fallbackNodes.length; ++i) {
                fallbackNodes[i].remove();
            }

        });

        window.alreadyProcessedMarkdeep = true;

        return;
    }
    
    // The following is Morgan's massive hack for allowing browsers to
    // directly parse Markdown from what appears to be a text file, but is
    // actually an intentionally malformed HTML file.
    
    // In order to be able to show what source files look like, the
    // noformat argument may be supplied.
    
    if (! noformat) {
        // Remove any recursive references to this script so that we don't trigger the cost of
        // recursive *loading*. (The alreadyProcessedMarkdeep variable will prevent recursive
        // *execution*.) We allow other scripts to pass through.
        toArray(document.getElementsByTagName('script')).forEach(function(node) {
            if (isMarkdeepScriptName(node.src)) {
                node.parentNode.removeChild(node);
            }
        });
        
        // Add an event handler for scrolling
        var scrollThreshold = parseInt(option('scrollThreshold'));
        document.addEventListener('scroll', function () {
            var b = document.body, c = b.classList, s = 'scrolled';
            if (b.scrollTop > scrollThreshold) c.add(s); else c.remove(s);
        });
        
        // Hide the body while formatting
        if (document.body) {
            document.body.style.visibility = 'hidden';
        }
    }
      
    var source = nodeToMarkdeepSource(document.body);

    if (noformat) { 
        // Abort processing. 
        source = source.rp(/<!-- Markdeep:.+$/gm, '') + MARKDEEP_LINE;
    
        // Escape the <> (not ampersand) that we just added
        source = source.rp(/</g, '&lt;').rp(/>/g, '&gt;');

        // Replace the Markdeep line itself so that ?noformat examples have a valid line to copy
        document.body.innerHTML = entag('pre', source);

        var fallbackNodes = document.getElementsByClassName('fallback');
        for (var i = 0; i < fallbackNodes.length; ++i) {
            fallbackNodes[i].remove();
        }

        return;
    }

    // In the common case of no INSERT commands, source is the original source
    // passed to avoid reparsing.
    var markdeepProcessor = function (source) {
        // Recompute the source text from the current version of the document
        // if it was unmodified
        source = source || nodeToMarkdeepSource(document.body);
        var markdeepHTML = markdeepToHTML(source, false);

        // console.log(markdeepHTML); // Final processed source 

        /////////////////////////////////////////////////////////////
        // Add the section header event handlers

        var onContextMenu = function (event) {
            var menu = null;
            try {
                // Test for whether the click was on a header
                var match = event.target.tagName.match(/^H(\d)$/);
                if (! match) { return; }

                // The event target is a header...ensure that it is a Markdeep header
                // (we could be in HTML or Doxygen mode and have non-.md content in the
                // same document)
                var node = event.target;
                while (node) {
                    if (node.classList.contains('md')) { break } else { node = node.parentElement; }
                }
                if (! node) {
                    // never found .md
                    return;
                }
                    
                // We are on a header
                var level = parseInt(match[1]) || 1;
                
                // Show the headerMenu
                menu = document.getElementById('mdContextMenu');
                if (! menu) { return; }
                
                var sectionType = ['Section', 'Subsection'][Math.min(level - 1, 1)];
                // Search backwards two siblings to grab the URL generated
                var anchorNode = event.target.previousElementSibling.previousElementSibling;
                
                var sectionName = event.target.innerText.trim();
                var sectionLabel = sectionName.toLowerCase();
                var anchor = anchorNode.name;
                var url = '' + location.origin + location.pathname + '#' + anchor;

                var shortUrl = url;
                if (shortUrl.length > 17) {
                    shortUrl = url.ss(0, 7) + '&hellip;' + location.pathname.ss(location.pathname.length - 8) + '#' + anchor;
                }
                
                var s = entag('div', 'Visit URL &ldquo;' + shortUrl + '&rdquo;',
                              'onclick="(location=&quot;' + url + '&quot;)"');
                
                s += entag('div', 'Copy URL &ldquo;' + shortUrl + '&rdquo;',
                           'onclick="navigator.clipboard.writeText(&quot;' + url + '&quot)&&(document.getElementById(\'mdContextMenu\').style.visibility=\'hidden\')"');
                
                s += entag('div', 'Copy Markdeep &ldquo;' + sectionName + ' ' + sectionType.toLowerCase() + '&rdquo;',
                           'onclick="navigator.clipboard.writeText(\'' + sectionName + ' ' + sectionType.toLowerCase() + '\')&&(document.getElementById(\'mdContextMenu\').style.visibility=\'hidden\')"');
                
                s += entag('div', 'Copy Markdeep &ldquo;' + sectionType + ' [' + sectionLabel + ']&rdquo;',
                           'onclick="navigator.clipboard.writeText(\'' + sectionType + ' [' + sectionLabel + ']\')&&(document.getElementById(\'mdContextMenu\').style.visibility=\'hidden\')"');
                
                s += entag('div', 'Copy HTML &ldquo;&lt;a href=&hellip;&gt;&rdquo;',
                           'onclick="navigator.clipboard.writeText(\'&lt;a href=&quot;' + url + '&quot;&gt;' + sectionName + '&lt;/a&gt;\')&&(document.getElementById(\'mdContextMenu\').style.visibility=\'hidden\')"');
                
                menu.innerHTML = s;
                menu.style.visibility = 'visible';
                menu.style.left = event.pageX + 'px';
                menu.style.top = event.pageY + 'px';
                
                event.preventDefault();
                return false;
            } catch (e) {
                // Something went wrong
                console.log(e);
                if (menu) { menu.style.visibility = 'hidden'; }
            }
        }

        markdeepHTML += '<div id="mdContextMenu" style="visibility:hidden"></div>';
        
        document.addEventListener('contextmenu', onContextMenu, false);
        document.addEventListener('mousedown', function (event) {
            var menu = document.getElementById('mdContextMenu');
            if (menu) {
                for (var node = event.target; node; node = node.parentElement) {
                    if (node === menu) { return; }
                }
                // Clicked off menu, so close it
                menu.style.visibility = 'hidden';
            }
        });
        document.addEventListener('keydown', function (event) {
            if (event.keyCode === 27) {
                var menu = document.getElementById('mdContextMenu');
                if (menu) { menu.style.visibility = 'hidden'; }
            }
        });
        
        
        /////////////////////////////////////////////////////////////
        
        var needMathJax = needsMathJax(markdeepHTML);
        if (needMathJax) {
            markdeepHTML = MATHJAX_CONFIG + markdeepHTML; 
        }
        
        markdeepHTML += MARKDEEP_FOOTER;
        
        // Replace the document. If using MathJax, include the custom Markdeep definitions
        var longDocument = source.length > 1000;
        
        // Setting "width" equal to 640 seems to give the best results on 
        // mobile devices in portrait mode. Setting "width=device-width" can cause markdeep
        // to appear exceedingly narrow on phones in the Chrome mobile preview.
        // https://developer.mozilla.org/en-US/docs/Mozilla/Mobile/Viewport_meta_tag
        var META = '<meta charset="UTF-8"><meta http-equiv="content-type" content="text/html;charset=UTF-8"><meta name="viewport" content="width=600, initial-scale=1">';
        var head = META + BODY_STYLESHEET + STYLESHEET + sectionNumberingStylesheet() + HIGHLIGHT_STYLESHEET;
        if (longDocument) {
            // Add more spacing before the title in a long document
            head += entag('style', 'div.title { padding-top: 40px; } div.afterTitles { height: 15px; }');
        }

        if (window.location.href.search(/\?.*export.*/i) !== -1) {
            // Export mode
            var text = head + document.head.innerHTML + markdeepHTML;
            if (needMathJax) {
                // Dynamically load mathjax
                text += '<script src="' + MATHJAX_URL +'"></script>';
            }
            document.body.innerHTML = entag('pre', escapeHTMLEntities(text));
        } else {
            document.head.innerHTML = head + document.head.innerHTML;
            document.body.innerHTML = markdeepHTML;
            if (needMathJax) { loadMathJax(); }            
        }

        // Change the ID of the body, so that CSS can distinguish Markdeep
        // controlling a whole document from Markdeep embedded within
        // a document in HTML mode.
        document.body.id = 'md';
        document.body.style.visibility = 'visible';

        var hashIndex = window.location.href.indexOf('#');
        if (hashIndex > -1) {
            // Scroll to the target; needed when loading is too fast (ironically)
            setTimeout(function () {
                var anchor = document.getElementsByName(window.location.href.substring(hashIndex + 1));
                if (anchor.length > 0) { anchor[0].scrollIntoView(); }
                if (window.markdeepOptions) (window.markdeepOptions.onLoad || Math.cos)();
            }, 100);
        } else if (window.markdeepOptions && window.markdeepOptions.onLoad) {
            // Wait for the DOM to update
            setTimeout(window.markdeepOptions.onLoad, 100);
        }
           
    };
    
    // Process insert commands, and then run the markdeepProcessor on the document
    processInsertCommands([document.body], [source], function (nodeArray, sourceArray) {
        markdeepProcessor(sourceArray && sourceArray[0]);
    });
}
    
})();

/*** Custom Highlight.js, only supports basic,vbnet,vbscript,csharp,javascript,python,json,xml,markdown, and plaintext */
/*!
  Highlight.js v11.5.1 (git: b8f233c8e2)
  (c) 2006-2022 Ivan Sagalaev and other contributors
  License: BSD-3-Clause
 */
var hljs=function(){"use strict";var e={exports:{}};function t(e){
return e instanceof Map?e.clear=e.delete=e.set=()=>{
throw Error("map is read-only")}:e instanceof Set&&(e.add=e.clear=e.delete=()=>{
throw Error("set is read-only")
}),Object.freeze(e),Object.getOwnPropertyNames(e).forEach((n=>{var i=e[n]
;"object"!=typeof i||Object.isFrozen(i)||t(i)})),e}
e.exports=t,e.exports.default=t;var n=e.exports;class i{constructor(e){
void 0===e.data&&(e.data={}),this.data=e.data,this.isMatchIgnored=!1}
ignoreMatch(){this.isMatchIgnored=!0}}function r(e){
return e.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#x27;")
}function s(e,...t){const n=Object.create(null);for(const t in e)n[t]=e[t]
;return t.forEach((e=>{for(const t in e)n[t]=e[t]})),n}const o=e=>!!e.kind
;class a{constructor(e,t){
this.buffer="",this.classPrefix=t.classPrefix,e.walk(this)}addText(e){
this.buffer+=r(e)}openNode(e){if(!o(e))return;let t=e.kind
;t=e.sublanguage?"language-"+t:((e,{prefix:t})=>{if(e.includes(".")){
const n=e.split(".")
;return[`${t}${n.shift()}`,...n.map(((e,t)=>`${e}${"_".repeat(t+1)}`))].join(" ")
}return`${t}${e}`})(t,{prefix:this.classPrefix}),this.span(t)}closeNode(e){
o(e)&&(this.buffer+="</span>")}value(){return this.buffer}span(e){
this.buffer+=`<span class="${e}">`}}class c{constructor(){this.rootNode={
children:[]},this.stack=[this.rootNode]}get top(){
return this.stack[this.stack.length-1]}get root(){return this.rootNode}add(e){
this.top.children.push(e)}openNode(e){const t={kind:e,children:[]}
;this.add(t),this.stack.push(t)}closeNode(){
if(this.stack.length>1)return this.stack.pop()}closeAllNodes(){
for(;this.closeNode(););}toJSON(){return JSON.stringify(this.rootNode,null,4)}
walk(e){return this.constructor._walk(e,this.rootNode)}static _walk(e,t){
return"string"==typeof t?e.addText(t):t.children&&(e.openNode(t),
t.children.forEach((t=>this._walk(e,t))),e.closeNode(t)),e}static _collapse(e){
"string"!=typeof e&&e.children&&(e.children.every((e=>"string"==typeof e))?e.children=[e.children.join("")]:e.children.forEach((e=>{
c._collapse(e)})))}}class l extends c{constructor(e){super(),this.options=e}
addKeyword(e,t){""!==e&&(this.openNode(t),this.addText(e),this.closeNode())}
addText(e){""!==e&&this.add(e)}addSublanguage(e,t){const n=e.root
;n.kind=t,n.sublanguage=!0,this.add(n)}toHTML(){
return new a(this,this.options).value()}finalize(){return!0}}function g(e){
return e?"string"==typeof e?e:e.source:null}function d(e){return f("(?=",e,")")}
function u(e){return f("(?:",e,")*")}function h(e){return f("(?:",e,")?")}
function f(...e){return e.map((e=>g(e))).join("")}function p(...e){const t=(e=>{
const t=e[e.length-1]
;return"object"==typeof t&&t.constructor===Object?(e.splice(e.length-1,1),t):{}
})(e);return"("+(t.capture?"":"?:")+e.map((e=>g(e))).join("|")+")"}
function b(e){return RegExp(e.toString()+"|").exec("").length-1}
const m=/\[(?:[^\\\]]|\\.)*\]|\(\??|\\([1-9][0-9]*)|\\./
;function E(e,{joinWith:t}){let n=0;return e.map((e=>{n+=1;const t=n
;let i=g(e),r="";for(;i.length>0;){const e=m.exec(i);if(!e){r+=i;break}
r+=i.substring(0,e.index),
i=i.substring(e.index+e[0].length),"\\"===e[0][0]&&e[1]?r+="\\"+(Number(e[1])+t):(r+=e[0],
"("===e[0]&&n++)}return r})).map((e=>`(${e})`)).join(t)}
const x="[a-zA-Z]\\w*",w="[a-zA-Z_]\\w*",y="\\b\\d+(\\.\\d+)?",_="(-?)(\\b0[xX][a-fA-F0-9]+|(\\b\\d+(\\.\\d*)?|\\.\\d+)([eE][-+]?\\d+)?)",k="\\b(0b[01]+)",v={
begin:"\\\\[\\s\\S]",relevance:0},O={scope:"string",begin:"'",end:"'",
illegal:"\\n",contains:[v]},N={scope:"string",begin:'"',end:'"',illegal:"\\n",
contains:[v]},M=(e,t,n={})=>{const i=s({scope:"comment",begin:e,end:t,
contains:[]},n);i.contains.push({scope:"doctag",
begin:"[ ]*(?=(TODO|FIXME|NOTE|BUG|OPTIMIZE|HACK|XXX):)",
end:/(TODO|FIXME|NOTE|BUG|OPTIMIZE|HACK|XXX):/,excludeBegin:!0,relevance:0})
;const r=p("I","a","is","so","us","to","at","if","in","it","on",/[A-Za-z]+['](d|ve|re|ll|t|s|n)/,/[A-Za-z]+[-][a-z]+/,/[A-Za-z][a-z]{2,}/)
;return i.contains.push({begin:f(/[ ]+/,"(",r,/[.]?[:]?([.][ ]|[ ])/,"){3}")}),i
},S=M("//","$"),R=M("/\\*","\\*/"),j=M("#","$");var A=Object.freeze({
__proto__:null,MATCH_NOTHING_RE:/\b\B/,IDENT_RE:x,UNDERSCORE_IDENT_RE:w,
NUMBER_RE:y,C_NUMBER_RE:_,BINARY_NUMBER_RE:k,
RE_STARTERS_RE:"!|!=|!==|%|%=|&|&&|&=|\\*|\\*=|\\+|\\+=|,|-|-=|/=|/|:|;|<<|<<=|<=|<|===|==|=|>>>=|>>=|>=|>>>|>>|>|\\?|\\[|\\{|\\(|\\^|\\^=|\\||\\|=|\\|\\||~",
SHEBANG:(e={})=>{const t=/^#![ ]*\//
;return e.binary&&(e.begin=f(t,/.*\b/,e.binary,/\b.*/)),s({scope:"meta",begin:t,
end:/$/,relevance:0,"on:begin":(e,t)=>{0!==e.index&&t.ignoreMatch()}},e)},
BACKSLASH_ESCAPE:v,APOS_STRING_MODE:O,QUOTE_STRING_MODE:N,PHRASAL_WORDS_MODE:{
begin:/\b(a|an|the|are|I'm|isn't|don't|doesn't|won't|but|just|should|pretty|simply|enough|gonna|going|wtf|so|such|will|you|your|they|like|more)\b/
},COMMENT:M,C_LINE_COMMENT_MODE:S,C_BLOCK_COMMENT_MODE:R,HASH_COMMENT_MODE:j,
NUMBER_MODE:{scope:"number",begin:y,relevance:0},C_NUMBER_MODE:{scope:"number",
begin:_,relevance:0},BINARY_NUMBER_MODE:{scope:"number",begin:k,relevance:0},
REGEXP_MODE:{begin:/(?=\/[^/\n]*\/)/,contains:[{scope:"regexp",begin:/\//,
end:/\/[gimuy]*/,illegal:/\n/,contains:[v,{begin:/\[/,end:/\]/,relevance:0,
contains:[v]}]}]},TITLE_MODE:{scope:"title",begin:x,relevance:0},
UNDERSCORE_TITLE_MODE:{scope:"title",begin:w,relevance:0},METHOD_GUARD:{
begin:"\\.\\s*[a-zA-Z_]\\w*",relevance:0},END_SAME_AS_BEGIN:e=>Object.assign(e,{
"on:begin":(e,t)=>{t.data._beginMatch=e[1]},"on:end":(e,t)=>{
t.data._beginMatch!==e[1]&&t.ignoreMatch()}})});function I(e,t){
"."===e.input[e.index-1]&&t.ignoreMatch()}function T(e,t){
void 0!==e.className&&(e.scope=e.className,delete e.className)}function L(e,t){
t&&e.beginKeywords&&(e.begin="\\b("+e.beginKeywords.split(" ").join("|")+")(?!\\.)(?=\\b|\\s)",
e.__beforeBegin=I,e.keywords=e.keywords||e.beginKeywords,delete e.beginKeywords,
void 0===e.relevance&&(e.relevance=0))}function B(e,t){
Array.isArray(e.illegal)&&(e.illegal=p(...e.illegal))}function D(e,t){
if(e.match){
if(e.begin||e.end)throw Error("begin & end are not supported with match")
;e.begin=e.match,delete e.match}}function H(e,t){
void 0===e.relevance&&(e.relevance=1)}const P=(e,t)=>{if(!e.beforeMatch)return
;if(e.starts)throw Error("beforeMatch cannot be used with starts")
;const n=Object.assign({},e);Object.keys(e).forEach((t=>{delete e[t]
})),e.keywords=n.keywords,e.begin=f(n.beforeMatch,d(n.begin)),e.starts={
relevance:0,contains:[Object.assign(n,{endsParent:!0})]
},e.relevance=0,delete n.beforeMatch
},C=["of","and","for","in","not","or","if","then","parent","list","value"]
;function $(e,t,n="keyword"){const i=Object.create(null)
;return"string"==typeof e?r(n,e.split(" ")):Array.isArray(e)?r(n,e):Object.keys(e).forEach((n=>{
Object.assign(i,$(e[n],t,n))})),i;function r(e,n){
t&&(n=n.map((e=>e.toLowerCase()))),n.forEach((t=>{const n=t.split("|")
;i[n[0]]=[e,U(n[0],n[1])]}))}}function U(e,t){
return t?Number(t):(e=>C.includes(e.toLowerCase()))(e)?0:1}const z={},K=e=>{
console.error(e)},W=(e,...t)=>{console.log("WARN: "+e,...t)},X=(e,t)=>{
z[`${e}/${t}`]||(console.log(`Deprecated as of ${e}. ${t}`),z[`${e}/${t}`]=!0)
},G=Error();function Z(e,t,{key:n}){let i=0;const r=e[n],s={},o={}
;for(let e=1;e<=t.length;e++)o[e+i]=r[e],s[e+i]=!0,i+=b(t[e-1])
;e[n]=o,e[n]._emit=s,e[n]._multi=!0}function F(e){(e=>{
e.scope&&"object"==typeof e.scope&&null!==e.scope&&(e.beginScope=e.scope,
delete e.scope)})(e),"string"==typeof e.beginScope&&(e.beginScope={
_wrap:e.beginScope}),"string"==typeof e.endScope&&(e.endScope={_wrap:e.endScope
}),(e=>{if(Array.isArray(e.begin)){
if(e.skip||e.excludeBegin||e.returnBegin)throw K("skip, excludeBegin, returnBegin not compatible with beginScope: {}"),
G
;if("object"!=typeof e.beginScope||null===e.beginScope)throw K("beginScope must be object"),
G;Z(e,e.begin,{key:"beginScope"}),e.begin=E(e.begin,{joinWith:""})}})(e),(e=>{
if(Array.isArray(e.end)){
if(e.skip||e.excludeEnd||e.returnEnd)throw K("skip, excludeEnd, returnEnd not compatible with endScope: {}"),
G
;if("object"!=typeof e.endScope||null===e.endScope)throw K("endScope must be object"),
G;Z(e,e.end,{key:"endScope"}),e.end=E(e.end,{joinWith:""})}})(e)}function V(e){
function t(t,n){
return RegExp(g(t),"m"+(e.case_insensitive?"i":"")+(e.unicodeRegex?"u":"")+(n?"g":""))
}class n{constructor(){
this.matchIndexes={},this.regexes=[],this.matchAt=1,this.position=0}
addRule(e,t){
t.position=this.position++,this.matchIndexes[this.matchAt]=t,this.regexes.push([t,e]),
this.matchAt+=b(e)+1}compile(){0===this.regexes.length&&(this.exec=()=>null)
;const e=this.regexes.map((e=>e[1]));this.matcherRe=t(E(e,{joinWith:"|"
}),!0),this.lastIndex=0}exec(e){this.matcherRe.lastIndex=this.lastIndex
;const t=this.matcherRe.exec(e);if(!t)return null
;const n=t.findIndex(((e,t)=>t>0&&void 0!==e)),i=this.matchIndexes[n]
;return t.splice(0,n),Object.assign(t,i)}}class i{constructor(){
this.rules=[],this.multiRegexes=[],
this.count=0,this.lastIndex=0,this.regexIndex=0}getMatcher(e){
if(this.multiRegexes[e])return this.multiRegexes[e];const t=new n
;return this.rules.slice(e).forEach((([e,n])=>t.addRule(e,n))),
t.compile(),this.multiRegexes[e]=t,t}resumingScanAtSamePosition(){
return 0!==this.regexIndex}considerAll(){this.regexIndex=0}addRule(e,t){
this.rules.push([e,t]),"begin"===t.type&&this.count++}exec(e){
const t=this.getMatcher(this.regexIndex);t.lastIndex=this.lastIndex
;let n=t.exec(e)
;if(this.resumingScanAtSamePosition())if(n&&n.index===this.lastIndex);else{
const t=this.getMatcher(0);t.lastIndex=this.lastIndex+1,n=t.exec(e)}
return n&&(this.regexIndex+=n.position+1,
this.regexIndex===this.count&&this.considerAll()),n}}
if(e.compilerExtensions||(e.compilerExtensions=[]),
e.contains&&e.contains.includes("self"))throw Error("ERR: contains `self` is not supported at the top-level of a language.  See documentation.")
;return e.classNameAliases=s(e.classNameAliases||{}),function n(r,o){const a=r
;if(r.isCompiled)return a
;[T,D,F,P].forEach((e=>e(r,o))),e.compilerExtensions.forEach((e=>e(r,o))),
r.__beforeBegin=null,[L,B,H].forEach((e=>e(r,o))),r.isCompiled=!0;let c=null
;return"object"==typeof r.keywords&&r.keywords.$pattern&&(r.keywords=Object.assign({},r.keywords),
c=r.keywords.$pattern,
delete r.keywords.$pattern),c=c||/\w+/,r.keywords&&(r.keywords=$(r.keywords,e.case_insensitive)),
a.keywordPatternRe=t(c,!0),
o&&(r.begin||(r.begin=/\B|\b/),a.beginRe=t(a.begin),r.end||r.endsWithParent||(r.end=/\B|\b/),
r.end&&(a.endRe=t(a.end)),
a.terminatorEnd=g(a.end)||"",r.endsWithParent&&o.terminatorEnd&&(a.terminatorEnd+=(r.end?"|":"")+o.terminatorEnd)),
r.illegal&&(a.illegalRe=t(r.illegal)),
r.contains||(r.contains=[]),r.contains=[].concat(...r.contains.map((e=>(e=>(e.variants&&!e.cachedVariants&&(e.cachedVariants=e.variants.map((t=>s(e,{
variants:null},t)))),e.cachedVariants?e.cachedVariants:q(e)?s(e,{
starts:e.starts?s(e.starts):null
}):Object.isFrozen(e)?s(e):e))("self"===e?r:e)))),r.contains.forEach((e=>{n(e,a)
})),r.starts&&n(r.starts,o),a.matcher=(e=>{const t=new i
;return e.contains.forEach((e=>t.addRule(e.begin,{rule:e,type:"begin"
}))),e.terminatorEnd&&t.addRule(e.terminatorEnd,{type:"end"
}),e.illegal&&t.addRule(e.illegal,{type:"illegal"}),t})(a),a}(e)}function q(e){
return!!e&&(e.endsWithParent||q(e.starts))}class J extends Error{
constructor(e,t){super(e),this.name="HTMLInjectionError",this.html=t}}
const Y=r,Q=s,ee=Symbol("nomatch");var te=(e=>{
const t=Object.create(null),r=Object.create(null),s=[];let o=!0
;const a="Could not find the language '{}', did you forget to load/include a language module?",c={
disableAutodetect:!0,name:"Plain text",contains:[]};let g={
ignoreUnescapedHTML:!1,throwUnescapedHTML:!1,noHighlightRe:/^(no-?highlight)$/i,
languageDetectRe:/\blang(?:uage)?-([\w-]+)\b/i,classPrefix:"hljs-",
cssSelector:"pre code",languages:null,__emitter:l};function b(e){
return g.noHighlightRe.test(e)}function m(e,t,n){let i="",r=""
;"object"==typeof t?(i=e,
n=t.ignoreIllegals,r=t.language):(X("10.7.0","highlight(lang, code, ...args) has been deprecated."),
X("10.7.0","Please use highlight(code, options) instead.\nhttps://github.com/highlightjs/highlight.js/issues/2277"),
r=e,i=t),void 0===n&&(n=!0);const s={code:i,language:r};N("before:highlight",s)
;const o=s.result?s.result:E(s.language,s.code,n)
;return o.code=s.code,N("after:highlight",o),o}function E(e,n,r,s){
const c=Object.create(null);function l(){if(!O.keywords)return void M.addText(S)
;let e=0;O.keywordPatternRe.lastIndex=0;let t=O.keywordPatternRe.exec(S),n=""
;for(;t;){n+=S.substring(e,t.index)
;const r=y.case_insensitive?t[0].toLowerCase():t[0],s=(i=r,O.keywords[i]);if(s){
const[e,i]=s
;if(M.addText(n),n="",c[r]=(c[r]||0)+1,c[r]<=7&&(R+=i),e.startsWith("_"))n+=t[0];else{
const n=y.classNameAliases[e]||e;M.addKeyword(t[0],n)}}else n+=t[0]
;e=O.keywordPatternRe.lastIndex,t=O.keywordPatternRe.exec(S)}var i
;n+=S.substr(e),M.addText(n)}function d(){null!=O.subLanguage?(()=>{
if(""===S)return;let e=null;if("string"==typeof O.subLanguage){
if(!t[O.subLanguage])return void M.addText(S)
;e=E(O.subLanguage,S,!0,N[O.subLanguage]),N[O.subLanguage]=e._top
}else e=x(S,O.subLanguage.length?O.subLanguage:null)
;O.relevance>0&&(R+=e.relevance),M.addSublanguage(e._emitter,e.language)
})():l(),S=""}function u(e,t){let n=1;const i=t.length-1;for(;n<=i;){
if(!e._emit[n]){n++;continue}const i=y.classNameAliases[e[n]]||e[n],r=t[n]
;i?M.addKeyword(r,i):(S=r,l(),S=""),n++}}function h(e,t){
return e.scope&&"string"==typeof e.scope&&M.openNode(y.classNameAliases[e.scope]||e.scope),
e.beginScope&&(e.beginScope._wrap?(M.addKeyword(S,y.classNameAliases[e.beginScope._wrap]||e.beginScope._wrap),
S=""):e.beginScope._multi&&(u(e.beginScope,t),S="")),O=Object.create(e,{parent:{
value:O}}),O}function f(e,t,n){let r=((e,t)=>{const n=e&&e.exec(t)
;return n&&0===n.index})(e.endRe,n);if(r){if(e["on:end"]){const n=new i(e)
;e["on:end"](t,n),n.isMatchIgnored&&(r=!1)}if(r){
for(;e.endsParent&&e.parent;)e=e.parent;return e}}
if(e.endsWithParent)return f(e.parent,t,n)}function p(e){
return 0===O.matcher.regexIndex?(S+=e[0],1):(I=!0,0)}function b(e){
const t=e[0],i=n.substr(e.index),r=f(O,e,i);if(!r)return ee;const s=O
;O.endScope&&O.endScope._wrap?(d(),
M.addKeyword(t,O.endScope._wrap)):O.endScope&&O.endScope._multi?(d(),
u(O.endScope,e)):s.skip?S+=t:(s.returnEnd||s.excludeEnd||(S+=t),
d(),s.excludeEnd&&(S=t));do{
O.scope&&M.closeNode(),O.skip||O.subLanguage||(R+=O.relevance),O=O.parent
}while(O!==r.parent);return r.starts&&h(r.starts,e),s.returnEnd?0:t.length}
let m={};function w(t,s){const a=s&&s[0];if(S+=t,null==a)return d(),0
;if("begin"===m.type&&"end"===s.type&&m.index===s.index&&""===a){
if(S+=n.slice(s.index,s.index+1),!o){const t=Error(`0 width match regex (${e})`)
;throw t.languageName=e,t.badRule=m.rule,t}return 1}
if(m=s,"begin"===s.type)return(e=>{
const t=e[0],n=e.rule,r=new i(n),s=[n.__beforeBegin,n["on:begin"]]
;for(const n of s)if(n&&(n(e,r),r.isMatchIgnored))return p(t)
;return n.skip?S+=t:(n.excludeBegin&&(S+=t),
d(),n.returnBegin||n.excludeBegin||(S=t)),h(n,e),n.returnBegin?0:t.length})(s)
;if("illegal"===s.type&&!r){
const e=Error('Illegal lexeme "'+a+'" for mode "'+(O.scope||"<unnamed>")+'"')
;throw e.mode=O,e}if("end"===s.type){const e=b(s);if(e!==ee)return e}
if("illegal"===s.type&&""===a)return 1
;if(A>1e5&&A>3*s.index)throw Error("potential infinite loop, way more iterations than matches")
;return S+=a,a.length}const y=k(e)
;if(!y)throw K(a.replace("{}",e)),Error('Unknown language: "'+e+'"')
;const _=V(y);let v="",O=s||_;const N={},M=new g.__emitter(g);(()=>{const e=[]
;for(let t=O;t!==y;t=t.parent)t.scope&&e.unshift(t.scope)
;e.forEach((e=>M.openNode(e)))})();let S="",R=0,j=0,A=0,I=!1;try{
for(O.matcher.considerAll();;){
A++,I?I=!1:O.matcher.considerAll(),O.matcher.lastIndex=j
;const e=O.matcher.exec(n);if(!e)break;const t=w(n.substring(j,e.index),e)
;j=e.index+t}return w(n.substr(j)),M.closeAllNodes(),M.finalize(),v=M.toHTML(),{
language:e,value:v,relevance:R,illegal:!1,_emitter:M,_top:O}}catch(t){
if(t.message&&t.message.includes("Illegal"))return{language:e,value:Y(n),
illegal:!0,relevance:0,_illegalBy:{message:t.message,index:j,
context:n.slice(j-100,j+100),mode:t.mode,resultSoFar:v},_emitter:M};if(o)return{
language:e,value:Y(n),illegal:!1,relevance:0,errorRaised:t,_emitter:M,_top:O}
;throw t}}function x(e,n){n=n||g.languages||Object.keys(t);const i=(e=>{
const t={value:Y(e),illegal:!1,relevance:0,_top:c,_emitter:new g.__emitter(g)}
;return t._emitter.addText(e),t})(e),r=n.filter(k).filter(O).map((t=>E(t,e,!1)))
;r.unshift(i);const s=r.sort(((e,t)=>{
if(e.relevance!==t.relevance)return t.relevance-e.relevance
;if(e.language&&t.language){if(k(e.language).supersetOf===t.language)return 1
;if(k(t.language).supersetOf===e.language)return-1}return 0})),[o,a]=s,l=o
;return l.secondBest=a,l}function w(e){let t=null;const n=(e=>{
let t=e.className+" ";t+=e.parentNode?e.parentNode.className:""
;const n=g.languageDetectRe.exec(t);if(n){const t=k(n[1])
;return t||(W(a.replace("{}",n[1])),
W("Falling back to no-highlight mode for this block.",e)),t?n[1]:"no-highlight"}
return t.split(/\s+/).find((e=>b(e)||k(e)))})(e);if(b(n))return
;if(N("before:highlightElement",{el:e,language:n
}),e.children.length>0&&(g.ignoreUnescapedHTML||(console.warn("One of your code blocks includes unescaped HTML. This is a potentially serious security risk."),
console.warn("https://github.com/highlightjs/highlight.js/wiki/security"),
console.warn("The element with unescaped HTML:"),
console.warn(e)),g.throwUnescapedHTML))throw new J("One of your code blocks includes unescaped HTML.",e.innerHTML)
;t=e;const i=t.textContent,s=n?m(i,{language:n,ignoreIllegals:!0}):x(i)
;e.innerHTML=s.value,((e,t,n)=>{const i=t&&r[t]||n
;e.classList.add("hljs"),e.classList.add("language-"+i)
})(e,n,s.language),e.result={language:s.language,re:s.relevance,
relevance:s.relevance},s.secondBest&&(e.secondBest={
language:s.secondBest.language,relevance:s.secondBest.relevance
}),N("after:highlightElement",{el:e,result:s,text:i})}let y=!1;function _(){
"loading"!==document.readyState?document.querySelectorAll(g.cssSelector).forEach(w):y=!0
}function k(e){return e=(e||"").toLowerCase(),t[e]||t[r[e]]}
function v(e,{languageName:t}){"string"==typeof e&&(e=[e]),e.forEach((e=>{
r[e.toLowerCase()]=t}))}function O(e){const t=k(e)
;return t&&!t.disableAutodetect}function N(e,t){const n=e;s.forEach((e=>{
e[n]&&e[n](t)}))}
"undefined"!=typeof window&&window.addEventListener&&window.addEventListener("DOMContentLoaded",(()=>{
y&&_()}),!1),Object.assign(e,{highlight:m,highlightAuto:x,highlightAll:_,
highlightElement:w,
highlightBlock:e=>(X("10.7.0","highlightBlock will be removed entirely in v12.0"),
X("10.7.0","Please use highlightElement now."),w(e)),configure:e=>{g=Q(g,e)},
initHighlighting:()=>{
_(),X("10.6.0","initHighlighting() deprecated.  Use highlightAll() now.")},
initHighlightingOnLoad:()=>{
_(),X("10.6.0","initHighlightingOnLoad() deprecated.  Use highlightAll() now.")
},registerLanguage:(n,i)=>{let r=null;try{r=i(e)}catch(e){
if(K("Language definition for '{}' could not be registered.".replace("{}",n)),
!o)throw e;K(e),r=c}
r.name||(r.name=n),t[n]=r,r.rawDefinition=i.bind(null,e),r.aliases&&v(r.aliases,{
languageName:n})},unregisterLanguage:e=>{delete t[e]
;for(const t of Object.keys(r))r[t]===e&&delete r[t]},
listLanguages:()=>Object.keys(t),getLanguage:k,registerAliases:v,
autoDetection:O,inherit:Q,addPlugin:e=>{(e=>{
e["before:highlightBlock"]&&!e["before:highlightElement"]&&(e["before:highlightElement"]=t=>{
e["before:highlightBlock"](Object.assign({block:t.el},t))
}),e["after:highlightBlock"]&&!e["after:highlightElement"]&&(e["after:highlightElement"]=t=>{
e["after:highlightBlock"](Object.assign({block:t.el},t))})})(e),s.push(e)}
}),e.debugMode=()=>{o=!1},e.safeMode=()=>{o=!0
},e.versionString="11.5.1",e.regex={concat:f,lookahead:d,either:p,optional:h,
anyNumberOfTimes:u};for(const e in A)"object"==typeof A[e]&&n(A[e])
;return Object.assign(e,A),e})({});return te}()
;"object"==typeof exports&&"undefined"!=typeof module&&(module.exports=hljs);/*! `python` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>{
const n=e.regex,a=/[\p{XID_Start}_]\p{XID_Continue}*/u,i=["and","as","assert","async","await","break","class","continue","def","del","elif","else","except","finally","for","from","global","if","import","in","is","lambda","nonlocal|10","not","or","pass","raise","return","try","while","with","yield"],s={
$pattern:/[A-Za-z]\w+|__\w+__/,keyword:i,
built_in:["__import__","abs","all","any","ascii","bin","bool","breakpoint","bytearray","bytes","callable","chr","classmethod","compile","complex","delattr","dict","dir","divmod","enumerate","eval","exec","filter","float","format","frozenset","getattr","globals","hasattr","hash","help","hex","id","input","int","isinstance","issubclass","iter","len","list","locals","map","max","memoryview","min","next","object","oct","open","ord","pow","print","property","range","repr","reversed","round","set","setattr","slice","sorted","staticmethod","str","sum","super","tuple","type","vars","zip"],
literal:["__debug__","Ellipsis","False","None","NotImplemented","True"],
type:["Any","Callable","Coroutine","Dict","List","Literal","Generic","Optional","Sequence","Set","Tuple","Type","Union"]
},t={className:"meta",begin:/^(>>>|\.\.\.) /},r={className:"subst",begin:/\{/,
end:/\}/,keywords:s,illegal:/#/},l={begin:/\{\{/,relevance:0},b={
className:"string",contains:[e.BACKSLASH_ESCAPE],variants:[{
begin:/([uU]|[bB]|[rR]|[bB][rR]|[rR][bB])?'''/,end:/'''/,
contains:[e.BACKSLASH_ESCAPE,t],relevance:10},{
begin:/([uU]|[bB]|[rR]|[bB][rR]|[rR][bB])?"""/,end:/"""/,
contains:[e.BACKSLASH_ESCAPE,t],relevance:10},{
begin:/([fF][rR]|[rR][fF]|[fF])'''/,end:/'''/,
contains:[e.BACKSLASH_ESCAPE,t,l,r]},{begin:/([fF][rR]|[rR][fF]|[fF])"""/,
end:/"""/,contains:[e.BACKSLASH_ESCAPE,t,l,r]},{begin:/([uU]|[rR])'/,end:/'/,
relevance:10},{begin:/([uU]|[rR])"/,end:/"/,relevance:10},{
begin:/([bB]|[bB][rR]|[rR][bB])'/,end:/'/},{begin:/([bB]|[bB][rR]|[rR][bB])"/,
end:/"/},{begin:/([fF][rR]|[rR][fF]|[fF])'/,end:/'/,
contains:[e.BACKSLASH_ESCAPE,l,r]},{begin:/([fF][rR]|[rR][fF]|[fF])"/,end:/"/,
contains:[e.BACKSLASH_ESCAPE,l,r]},e.APOS_STRING_MODE,e.QUOTE_STRING_MODE]
},o="[0-9](_?[0-9])*",c=`(\\b(${o}))?\\.(${o})|\\b(${o})\\.`,d="\\b|"+i.join("|"),g={
className:"number",relevance:0,variants:[{
begin:`(\\b(${o})|(${c}))[eE][+-]?(${o})[jJ]?(?=${d})`},{begin:`(${c})[jJ]?`},{
begin:`\\b([1-9](_?[0-9])*|0+(_?0)*)[lLjJ]?(?=${d})`},{
begin:`\\b0[bB](_?[01])+[lL]?(?=${d})`},{begin:`\\b0[oO](_?[0-7])+[lL]?(?=${d})`
},{begin:`\\b0[xX](_?[0-9a-fA-F])+[lL]?(?=${d})`},{begin:`\\b(${o})[jJ](?=${d})`
}]},p={className:"comment",begin:n.lookahead(/# type:/),end:/$/,keywords:s,
contains:[{begin:/# type:/},{begin:/#/,end:/\b\B/,endsWithParent:!0}]},m={
className:"params",variants:[{className:"",begin:/\(\s*\)/,skip:!0},{begin:/\(/,
end:/\)/,excludeBegin:!0,excludeEnd:!0,keywords:s,
contains:["self",t,g,b,e.HASH_COMMENT_MODE]}]};return r.contains=[b,g,t],{
name:"Python",aliases:["py","gyp","ipython"],unicodeRegex:!0,keywords:s,
illegal:/(<\/|->|\?)|=>/,contains:[t,g,{begin:/\bself\b/},{beginKeywords:"if",
relevance:0},b,p,e.HASH_COMMENT_MODE,{match:[/\bdef/,/\s+/,a],scope:{
1:"keyword",3:"title.function"},contains:[m]},{variants:[{
match:[/\bclass/,/\s+/,a,/\s*/,/\(\s*/,a,/\s*\)/]},{match:[/\bclass/,/\s+/,a]}],
scope:{1:"keyword",3:"title.class",6:"title.class.inherited"}},{
className:"meta",begin:/^[\t ]*@/,end:/(?=#)|$/,contains:[g,m,b]}]}}})()
;hljs.registerLanguage("python",e)})();/*! `xml` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>{
const a=e.regex,n=a.concat(/[A-Z_]/,a.optional(/[A-Z0-9_.-]*:/),/[A-Z0-9_.-]*/),s={
className:"symbol",begin:/&[a-z]+;|&#[0-9]+;|&#x[a-f0-9]+;/},t={begin:/\s/,
contains:[{className:"keyword",begin:/#?[a-z_][a-z1-9_-]+/,illegal:/\n/}]
},i=e.inherit(t,{begin:/\(/,end:/\)/}),c=e.inherit(e.APOS_STRING_MODE,{
className:"string"}),l=e.inherit(e.QUOTE_STRING_MODE,{className:"string"}),r={
endsWithParent:!0,illegal:/</,relevance:0,contains:[{className:"attr",
begin:/[A-Za-z0-9._:-]+/,relevance:0},{begin:/=\s*/,relevance:0,contains:[{
className:"string",endsParent:!0,variants:[{begin:/"/,end:/"/,contains:[s]},{
begin:/'/,end:/'/,contains:[s]},{begin:/[^\s"'=<>`]+/}]}]}]};return{
name:"HTML, XML",
aliases:["html","xhtml","rss","atom","xjb","xsd","xsl","plist","wsf","svg"],
case_insensitive:!0,contains:[{className:"meta",begin:/<![a-z]/,end:/>/,
relevance:10,contains:[t,l,c,i,{begin:/\[/,end:/\]/,contains:[{className:"meta",
begin:/<![a-z]/,end:/>/,contains:[t,i,l,c]}]}]},e.COMMENT(/<!--/,/-->/,{
relevance:10}),{begin:/<!\[CDATA\[/,end:/\]\]>/,relevance:10},s,{
className:"meta",end:/\?>/,variants:[{begin:/<\?xml/,relevance:10,contains:[l]
},{begin:/<\?[a-z][a-z0-9]+/}]},{className:"tag",begin:/<style(?=\s|>)/,end:/>/,
keywords:{name:"style"},contains:[r],starts:{end:/<\/style>/,returnEnd:!0,
subLanguage:["css","xml"]}},{className:"tag",begin:/<script(?=\s|>)/,end:/>/,
keywords:{name:"script"},contains:[r],starts:{end:/<\/script>/,returnEnd:!0,
subLanguage:["javascript","handlebars","xml"]}},{className:"tag",begin:/<>|<\/>/
},{className:"tag",
begin:a.concat(/</,a.lookahead(a.concat(n,a.either(/\/>/,/>/,/\s/)))),
end:/\/?>/,contains:[{className:"name",begin:n,relevance:0,starts:r}]},{
className:"tag",begin:a.concat(/<\//,a.lookahead(a.concat(n,/>/))),contains:[{
className:"name",begin:n,relevance:0},{begin:/>/,relevance:0,endsParent:!0}]}]}}
})();hljs.registerLanguage("xml",e)})();/*! `json` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>({name:"JSON",contains:[{
className:"attr",begin:/"(\\.|[^\\"\r\n])*"(?=\s*:)/,relevance:1.01},{
match:/[{}[\],:]/,className:"punctuation",relevance:0},e.QUOTE_STRING_MODE,{
beginKeywords:"true false null"
},e.C_NUMBER_MODE,e.C_LINE_COMMENT_MODE,e.C_BLOCK_COMMENT_MODE],illegal:"\\S"})
})();hljs.registerLanguage("json",e)})();/*! `vbscript` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>{
const t=e.regex,r=["lcase","month","vartype","instrrev","ubound","setlocale","getobject","rgb","getref","string","weekdayname","rnd","dateadd","monthname","now","day","minute","isarray","cbool","round","formatcurrency","conversions","csng","timevalue","second","year","space","abs","clng","timeserial","fixs","len","asc","isempty","maths","dateserial","atn","timer","isobject","filter","weekday","datevalue","ccur","isdate","instr","datediff","formatdatetime","replace","isnull","right","sgn","array","snumeric","log","cdbl","hex","chr","lbound","msgbox","ucase","getlocale","cos","cdate","cbyte","rtrim","join","hour","oct","typename","trim","strcomp","int","createobject","loadpicture","tan","formatnumber","mid","split","cint","sin","datepart","ltrim","sqr","time","derived","eval","date","formatpercent","exp","inputbox","left","ascw","chrw","regexp","cstr","err"]
;return{name:"VBScript",aliases:["vbs"],case_insensitive:!0,keywords:{
keyword:["call","class","const","dim","do","loop","erase","execute","executeglobal","exit","for","each","next","function","if","then","else","on","error","option","explicit","new","private","property","let","get","public","randomize","redim","rem","select","case","set","stop","sub","while","wend","with","end","to","elseif","is","or","xor","and","not","class_initialize","class_terminate","default","preserve","in","me","byval","byref","step","resume","goto"],
built_in:["server","response","request","scriptengine","scriptenginebuildversion","scriptengineminorversion","scriptenginemajorversion"],
literal:["true","false","null","nothing","empty"]},illegal:"//",contains:[{
begin:t.concat(t.either(...r),"\\s*\\("),relevance:0,keywords:{built_in:r}
},e.inherit(e.QUOTE_STRING_MODE,{contains:[{begin:'""'}]}),e.COMMENT(/'/,/$/,{
relevance:0}),e.C_NUMBER_MODE]}}})();hljs.registerLanguage("vbscript",e)})();/*! `plaintext` grammar compiled for Highlight.js 11.5.1 */
(()=>{var t=(()=>{"use strict";return t=>({name:"Plain text",
aliases:["text","txt"],disableAutodetect:!0})})()
;hljs.registerLanguage("plaintext",t)})();/*! `vbnet` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>{
const n=e.regex,t=/\d{1,2}\/\d{1,2}\/\d{4}/,a=/\d{4}-\d{1,2}-\d{1,2}/,i=/(\d|1[012])(:\d+){0,2} *(AM|PM)/,s=/\d{1,2}(:\d{1,2}){1,2}/,r={
className:"literal",variants:[{begin:n.concat(/# */,n.either(a,t),/ *#/)},{
begin:n.concat(/# */,s,/ *#/)},{begin:n.concat(/# */,i,/ *#/)},{
begin:n.concat(/# */,n.either(a,t),/ +/,n.either(i,s),/ *#/)}]
},l=e.COMMENT(/'''/,/$/,{contains:[{className:"doctag",begin:/<\/?/,end:/>/}]
}),o=e.COMMENT(null,/$/,{variants:[{begin:/'/},{begin:/([\t ]|^)REM(?=\s)/}]})
;return{name:"Visual Basic .NET",aliases:["vb"],case_insensitive:!0,
classNameAliases:{label:"symbol"},keywords:{
keyword:"addhandler alias aggregate ansi as async assembly auto binary by byref byval call case catch class compare const continue custom declare default delegate dim distinct do each equals else elseif end enum erase error event exit explicit finally for friend from function get global goto group handles if implements imports in inherits interface into iterator join key let lib loop me mid module mustinherit mustoverride mybase myclass namespace narrowing new next notinheritable notoverridable of off on operator option optional order overloads overridable overrides paramarray partial preserve private property protected public raiseevent readonly redim removehandler resume return select set shadows shared skip static step stop structure strict sub synclock take text then throw to try unicode until using when where while widening with withevents writeonly yield",
built_in:"addressof and andalso await directcast gettype getxmlnamespace is isfalse isnot istrue like mod nameof new not or orelse trycast typeof xor cbool cbyte cchar cdate cdbl cdec cint clng cobj csbyte cshort csng cstr cuint culng cushort",
type:"boolean byte char date decimal double integer long object sbyte short single string uinteger ulong ushort",
literal:"true false nothing"},
illegal:"//|\\{|\\}|endif|gosub|variant|wend|^\\$ ",contains:[{
className:"string",begin:/"(""|[^/n])"C\b/},{className:"string",begin:/"/,
end:/"/,illegal:/\n/,contains:[{begin:/""/}]},r,{className:"number",relevance:0,
variants:[{begin:/\b\d[\d_]*((\.[\d_]+(E[+-]?[\d_]+)?)|(E[+-]?[\d_]+))[RFD@!#]?/
},{begin:/\b\d[\d_]*((U?[SIL])|[%&])?/},{begin:/&H[\dA-F_]+((U?[SIL])|[%&])?/},{
begin:/&O[0-7_]+((U?[SIL])|[%&])?/},{begin:/&B[01_]+((U?[SIL])|[%&])?/}]},{
className:"label",begin:/^\w+:/},l,o,{className:"meta",
begin:/[\t ]*#(const|disable|else|elseif|enable|end|externalsource|if|region)\b/,
end:/$/,keywords:{
keyword:"const disable else elseif enable end externalsource if region then"},
contains:[o]}]}}})();hljs.registerLanguage("vbnet",e)})();/*! `basic` grammar compiled for Highlight.js 11.5.1 */
(()=>{var E=(()=>{"use strict";return E=>({name:"BASIC",case_insensitive:!0,
illegal:"^.",keywords:{$pattern:"[a-zA-Z][a-zA-Z0-9_$%!#]*",
keyword:["ABS","ASC","AND","ATN","AUTO|0","BEEP","BLOAD|10","BSAVE|10","CALL","CALLS","CDBL","CHAIN","CHDIR","CHR$|10","CINT","CIRCLE","CLEAR","CLOSE","CLS","COLOR","COM","COMMON","CONT","COS","CSNG","CSRLIN","CVD","CVI","CVS","DATA","DATE$","DEFDBL","DEFINT","DEFSNG","DEFSTR","DEF|0","SEG","USR","DELETE","DIM","DRAW","EDIT","END","ENVIRON","ENVIRON$","EOF","EQV","ERASE","ERDEV","ERDEV$","ERL","ERR","ERROR","EXP","FIELD","FILES","FIX","FOR|0","FRE","GET","GOSUB|10","GOTO","HEX$","IF","THEN","ELSE|0","INKEY$","INP","INPUT","INPUT#","INPUT$","INSTR","IMP","INT","IOCTL","IOCTL$","KEY","ON","OFF","LIST","KILL","LEFT$","LEN","LET","LINE","LLIST","LOAD","LOC","LOCATE","LOF","LOG","LPRINT","USING","LSET","MERGE","MID$","MKDIR","MKD$","MKI$","MKS$","MOD","NAME","NEW","NEXT","NOISE","NOT","OCT$","ON","OR","PEN","PLAY","STRIG","OPEN","OPTION","BASE","OUT","PAINT","PALETTE","PCOPY","PEEK","PMAP","POINT","POKE","POS","PRINT","PRINT]","PSET","PRESET","PUT","RANDOMIZE","READ","REM","RENUM","RESET|0","RESTORE","RESUME","RETURN|0","RIGHT$","RMDIR","RND","RSET","RUN","SAVE","SCREEN","SGN","SHELL","SIN","SOUND","SPACE$","SPC","SQR","STEP","STICK","STOP","STR$","STRING$","SWAP","SYSTEM","TAB","TAN","TIME$","TIMER","TROFF","TRON","TO","USR","VAL","VARPTR","VARPTR$","VIEW","WAIT","WHILE","WEND","WIDTH","WINDOW","WRITE","XOR"]
},contains:[E.QUOTE_STRING_MODE,E.COMMENT("REM","$",{relevance:10
}),E.COMMENT("'","$",{relevance:0}),{className:"symbol",begin:"^[0-9]+ ",
relevance:10},{className:"number",begin:"\\b\\d+(\\.\\d+)?([edED]\\d+)?[#!]?",
relevance:0},{className:"number",begin:"(&[hH][0-9a-fA-F]{1,4})"},{
className:"number",begin:"(&[oO][0-7]{1,6})"}]})})()
;hljs.registerLanguage("basic",E)})();/*! `markdown` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>{const n={begin:/<\/?[A-Za-z_]/,
end:">",subLanguage:"xml",relevance:0},a={variants:[{begin:/\[.+?\]\[.*?\]/,
relevance:0},{
begin:/\[.+?\]\(((data|javascript|mailto):|(?:http|ftp)s?:\/\/).*?\)/,
relevance:2},{
begin:e.regex.concat(/\[.+?\]\(/,/[A-Za-z][A-Za-z0-9+.-]*/,/:\/\/.*?\)/),
relevance:2},{begin:/\[.+?\]\([./?&#].*?\)/,relevance:1},{
begin:/\[.*?\]\(.*?\)/,relevance:0}],returnBegin:!0,contains:[{match:/\[(?=\])/
},{className:"string",relevance:0,begin:"\\[",end:"\\]",excludeBegin:!0,
returnEnd:!0},{className:"link",relevance:0,begin:"\\]\\(",end:"\\)",
excludeBegin:!0,excludeEnd:!0},{className:"symbol",relevance:0,begin:"\\]\\[",
end:"\\]",excludeBegin:!0,excludeEnd:!0}]},i={className:"strong",contains:[],
variants:[{begin:/_{2}/,end:/_{2}/},{begin:/\*{2}/,end:/\*{2}/}]},s={
className:"emphasis",contains:[],variants:[{begin:/\*(?!\*)/,end:/\*/},{
begin:/_(?!_)/,end:/_/,relevance:0}]},c=e.inherit(i,{contains:[]
}),t=e.inherit(s,{contains:[]});i.contains.push(t),s.contains.push(c)
;let g=[n,a];return[i,s,c,t].forEach((e=>{e.contains=e.contains.concat(g)
})),g=g.concat(i,s),{name:"Markdown",aliases:["md","mkdown","mkd"],contains:[{
className:"section",variants:[{begin:"^#{1,6}",end:"$",contains:g},{
begin:"(?=^.+?\\n[=-]{2,}$)",contains:[{begin:"^[=-]*$"},{begin:"^",end:"\\n",
contains:g}]}]},n,{className:"bullet",begin:"^[ \t]*([*+-]|(\\d+\\.))(?=\\s+)",
end:"\\s+",excludeEnd:!0},i,s,{className:"quote",begin:"^>\\s+",contains:g,
end:"$"},{className:"code",variants:[{begin:"(`{3,})[^`](.|\\n)*?\\1`*[ ]*"},{
begin:"(~{3,})[^~](.|\\n)*?\\1~*[ ]*"},{begin:"```",end:"```+[ ]*$"},{
begin:"~~~",end:"~~~+[ ]*$"},{begin:"`.+?`"},{begin:"(?=^( {4}|\\t))",
contains:[{begin:"^( {4}|\\t)",end:"(\\n)$"}],relevance:0}]},{
begin:"^[-\\*]{3,}",end:"$"},a,{begin:/^\[[^\n]+\]:/,returnBegin:!0,contains:[{
className:"symbol",begin:/\[/,end:/\]/,excludeBegin:!0,excludeEnd:!0},{
className:"link",begin:/:\s*/,end:/$/,excludeBegin:!0}]}]}}})()
;hljs.registerLanguage("markdown",e)})();/*! `javascript` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict"
;const e="[A-Za-z$_][0-9A-Za-z$_]*",n=["as","in","of","if","for","while","finally","var","new","function","do","return","void","else","break","catch","instanceof","with","throw","case","default","try","switch","continue","typeof","delete","let","yield","const","class","debugger","async","await","static","import","from","export","extends"],a=["true","false","null","undefined","NaN","Infinity"],t=["Object","Function","Boolean","Symbol","Math","Date","Number","BigInt","String","RegExp","Array","Float32Array","Float64Array","Int8Array","Uint8Array","Uint8ClampedArray","Int16Array","Int32Array","Uint16Array","Uint32Array","BigInt64Array","BigUint64Array","Set","Map","WeakSet","WeakMap","ArrayBuffer","SharedArrayBuffer","Atomics","DataView","JSON","Promise","Generator","GeneratorFunction","AsyncFunction","Reflect","Proxy","Intl","WebAssembly"],s=["Error","EvalError","InternalError","RangeError","ReferenceError","SyntaxError","TypeError","URIError"],r=["setInterval","setTimeout","clearInterval","clearTimeout","require","exports","eval","isFinite","isNaN","parseFloat","parseInt","decodeURI","decodeURIComponent","encodeURI","encodeURIComponent","escape","unescape"],c=["arguments","this","super","console","window","document","localStorage","module","global"],i=[].concat(r,t,s)
;return o=>{const l=o.regex,b=e,d={begin:/<[A-Za-z0-9\\._:-]+/,
end:/\/[A-Za-z0-9\\._:-]+>|\/>/,isTrulyOpeningTag:(e,n)=>{
const a=e[0].length+e.index,t=e.input[a]
;if("<"===t||","===t)return void n.ignoreMatch();let s
;">"===t&&(((e,{after:n})=>{const a="</"+e[0].slice(1)
;return-1!==e.input.indexOf(a,n)})(e,{after:a
})||n.ignoreMatch()),(s=e.input.substr(a).match(/^\s+extends\s+/))&&0===s.index&&n.ignoreMatch()
}},g={$pattern:e,keyword:n,literal:a,built_in:i,"variable.language":c
},u="\\.([0-9](_?[0-9])*)",m="0|[1-9](_?[0-9])*|0[0-7]*[89][0-9]*",E={
className:"number",variants:[{
begin:`(\\b(${m})((${u})|\\.)?|(${u}))[eE][+-]?([0-9](_?[0-9])*)\\b`},{
begin:`\\b(${m})\\b((${u})\\b|\\.)?|(${u})\\b`},{
begin:"\\b(0|[1-9](_?[0-9])*)n\\b"},{
begin:"\\b0[xX][0-9a-fA-F](_?[0-9a-fA-F])*n?\\b"},{
begin:"\\b0[bB][0-1](_?[0-1])*n?\\b"},{begin:"\\b0[oO][0-7](_?[0-7])*n?\\b"},{
begin:"\\b0[0-7]+n?\\b"}],relevance:0},A={className:"subst",begin:"\\$\\{",
end:"\\}",keywords:g,contains:[]},y={begin:"html`",end:"",starts:{end:"`",
returnEnd:!1,contains:[o.BACKSLASH_ESCAPE,A],subLanguage:"xml"}},N={
begin:"css`",end:"",starts:{end:"`",returnEnd:!1,
contains:[o.BACKSLASH_ESCAPE,A],subLanguage:"css"}},_={className:"string",
begin:"`",end:"`",contains:[o.BACKSLASH_ESCAPE,A]},f={className:"comment",
variants:[o.COMMENT(/\/\*\*(?!\/)/,"\\*/",{relevance:0,contains:[{
begin:"(?=@[A-Za-z]+)",relevance:0,contains:[{className:"doctag",
begin:"@[A-Za-z]+"},{className:"type",begin:"\\{",end:"\\}",excludeEnd:!0,
excludeBegin:!0,relevance:0},{className:"variable",begin:b+"(?=\\s*(-)|$)",
endsParent:!0,relevance:0},{begin:/(?=[^\n])\s/,relevance:0}]}]
}),o.C_BLOCK_COMMENT_MODE,o.C_LINE_COMMENT_MODE]
},h=[o.APOS_STRING_MODE,o.QUOTE_STRING_MODE,y,N,_,E];A.contains=h.concat({
begin:/\{/,end:/\}/,keywords:g,contains:["self"].concat(h)})
;const v=[].concat(f,A.contains),p=v.concat([{begin:/\(/,end:/\)/,keywords:g,
contains:["self"].concat(v)}]),S={className:"params",begin:/\(/,end:/\)/,
excludeBegin:!0,excludeEnd:!0,keywords:g,contains:p},w={variants:[{
match:[/class/,/\s+/,b,/\s+/,/extends/,/\s+/,l.concat(b,"(",l.concat(/\./,b),")*")],
scope:{1:"keyword",3:"title.class",5:"keyword",7:"title.class.inherited"}},{
match:[/class/,/\s+/,b],scope:{1:"keyword",3:"title.class"}}]},R={relevance:0,
match:l.either(/\bJSON/,/\b[A-Z][a-z]+([A-Z][a-z]*|\d)*/,/\b[A-Z]{2,}([A-Z][a-z]+|\d)+([A-Z][a-z]*)*/,/\b[A-Z]{2,}[a-z]+([A-Z][a-z]+|\d)*([A-Z][a-z]*)*/),
className:"title.class",keywords:{_:[...t,...s]}},O={variants:[{
match:[/function/,/\s+/,b,/(?=\s*\()/]},{match:[/function/,/\s*(?=\()/]}],
className:{1:"keyword",3:"title.function"},label:"func.def",contains:[S],
illegal:/%/},k={
match:l.concat(/\b/,(I=[...r,"super"],l.concat("(?!",I.join("|"),")")),b,l.lookahead(/\(/)),
className:"title.function",relevance:0};var I;const x={
begin:l.concat(/\./,l.lookahead(l.concat(b,/(?![0-9A-Za-z$_(])/))),end:b,
excludeBegin:!0,keywords:"prototype",className:"property",relevance:0},T={
match:[/get|set/,/\s+/,b,/(?=\()/],className:{1:"keyword",3:"title.function"},
contains:[{begin:/\(\)/},S]
},C="(\\([^()]*(\\([^()]*(\\([^()]*\\)[^()]*)*\\)[^()]*)*\\)|"+o.UNDERSCORE_IDENT_RE+")\\s*=>",M={
match:[/const|var|let/,/\s+/,b,/\s*/,/=\s*/,/(async\s*)?/,l.lookahead(C)],
keywords:"async",className:{1:"keyword",3:"title.function"},contains:[S]}
;return{name:"Javascript",aliases:["js","jsx","mjs","cjs"],keywords:g,exports:{
PARAMS_CONTAINS:p,CLASS_REFERENCE:R},illegal:/#(?![$_A-z])/,
contains:[o.SHEBANG({label:"shebang",binary:"node",relevance:5}),{
label:"use_strict",className:"meta",relevance:10,
begin:/^\s*['"]use (strict|asm)['"]/
},o.APOS_STRING_MODE,o.QUOTE_STRING_MODE,y,N,_,f,E,R,{className:"attr",
begin:b+l.lookahead(":"),relevance:0},M,{
begin:"("+o.RE_STARTERS_RE+"|\\b(case|return|throw)\\b)\\s*",
keywords:"return throw case",relevance:0,contains:[f,o.REGEXP_MODE,{
className:"function",begin:C,returnBegin:!0,end:"\\s*=>",contains:[{
className:"params",variants:[{begin:o.UNDERSCORE_IDENT_RE,relevance:0},{
className:null,begin:/\(\s*\)/,skip:!0},{begin:/\(/,end:/\)/,excludeBegin:!0,
excludeEnd:!0,keywords:g,contains:p}]}]},{begin:/,/,relevance:0},{match:/\s+/,
relevance:0},{variants:[{begin:"<>",end:"</>"},{
match:/<[A-Za-z0-9\\._:-]+\s*\/>/},{begin:d.begin,
"on:begin":d.isTrulyOpeningTag,end:d.end}],subLanguage:"xml",contains:[{
begin:d.begin,end:d.end,skip:!0,contains:["self"]}]}]},O,{
beginKeywords:"while if switch catch for"},{
begin:"\\b(?!function)"+o.UNDERSCORE_IDENT_RE+"\\([^()]*(\\([^()]*(\\([^()]*\\)[^()]*)*\\)[^()]*)*\\)\\s*\\{",
returnBegin:!0,label:"func.def",contains:[S,o.inherit(o.TITLE_MODE,{begin:b,
className:"title.function"})]},{match:/\.\.\./,relevance:0},x,{match:"\\$"+b,
relevance:0},{match:[/\bconstructor(?=\s*\()/],className:{1:"title.function"},
contains:[S]},k,{relevance:0,match:/\b[A-Z][A-Z_0-9]+\b/,
className:"variable.constant"},w,T,{match:/\$[(.]/}]}}})()
;hljs.registerLanguage("javascript",e)})();/*! `csharp` grammar compiled for Highlight.js 11.5.1 */
(()=>{var e=(()=>{"use strict";return e=>{const n={
keyword:["abstract","as","base","break","case","catch","class","const","continue","do","else","event","explicit","extern","finally","fixed","for","foreach","goto","if","implicit","in","interface","internal","is","lock","namespace","new","operator","out","override","params","private","protected","public","readonly","record","ref","return","sealed","sizeof","stackalloc","static","struct","switch","this","throw","try","typeof","unchecked","unsafe","using","virtual","void","volatile","while"].concat(["add","alias","and","ascending","async","await","by","descending","equals","from","get","global","group","init","into","join","let","nameof","not","notnull","on","or","orderby","partial","remove","select","set","unmanaged","value|0","var","when","where","with","yield"]),
built_in:["bool","byte","char","decimal","delegate","double","dynamic","enum","float","int","long","nint","nuint","object","sbyte","short","string","ulong","uint","ushort"],
literal:["default","false","null","true"]},a=e.inherit(e.TITLE_MODE,{
begin:"[a-zA-Z](\\.?\\w)*"}),i={className:"number",variants:[{
begin:"\\b(0b[01']+)"},{
begin:"(-?)\\b([\\d']+(\\.[\\d']*)?|\\.[\\d']+)(u|U|l|L|ul|UL|f|F|b|B)"},{
begin:"(-?)(\\b0[xX][a-fA-F0-9']+|(\\b[\\d']+(\\.[\\d']*)?|\\.[\\d']+)([eE][-+]?[\\d']+)?)"
}],relevance:0},s={className:"string",begin:'@"',end:'"',contains:[{begin:'""'}]
},t=e.inherit(s,{illegal:/\n/}),r={className:"subst",begin:/\{/,end:/\}/,
keywords:n},l=e.inherit(r,{illegal:/\n/}),c={className:"string",begin:/\$"/,
end:'"',illegal:/\n/,contains:[{begin:/\{\{/},{begin:/\}\}/
},e.BACKSLASH_ESCAPE,l]},o={className:"string",begin:/\$@"/,end:'"',contains:[{
begin:/\{\{/},{begin:/\}\}/},{begin:'""'},r]},d=e.inherit(o,{illegal:/\n/,
contains:[{begin:/\{\{/},{begin:/\}\}/},{begin:'""'},l]})
;r.contains=[o,c,s,e.APOS_STRING_MODE,e.QUOTE_STRING_MODE,i,e.C_BLOCK_COMMENT_MODE],
l.contains=[d,c,t,e.APOS_STRING_MODE,e.QUOTE_STRING_MODE,i,e.inherit(e.C_BLOCK_COMMENT_MODE,{
illegal:/\n/})];const g={variants:[o,c,s,e.APOS_STRING_MODE,e.QUOTE_STRING_MODE]
},E={begin:"<",end:">",contains:[{beginKeywords:"in out"},a]
},_=e.IDENT_RE+"(<"+e.IDENT_RE+"(\\s*,\\s*"+e.IDENT_RE+")*>)?(\\[\\])?",b={
begin:"@"+e.IDENT_RE,relevance:0};return{name:"C#",aliases:["cs","c#"],
keywords:n,illegal:/::/,contains:[e.COMMENT("///","$",{returnBegin:!0,
contains:[{className:"doctag",variants:[{begin:"///",relevance:0},{
begin:"\x3c!--|--\x3e"},{begin:"</?",end:">"}]}]
}),e.C_LINE_COMMENT_MODE,e.C_BLOCK_COMMENT_MODE,{className:"meta",begin:"#",
end:"$",keywords:{
keyword:"if else elif endif define undef warning error line region endregion pragma checksum"
}},g,i,{beginKeywords:"class interface",relevance:0,end:/[{;=]/,
illegal:/[^\s:,]/,contains:[{beginKeywords:"where class"
},a,E,e.C_LINE_COMMENT_MODE,e.C_BLOCK_COMMENT_MODE]},{beginKeywords:"namespace",
relevance:0,end:/[{;=]/,illegal:/[^\s:]/,
contains:[a,e.C_LINE_COMMENT_MODE,e.C_BLOCK_COMMENT_MODE]},{
beginKeywords:"record",relevance:0,end:/[{;=]/,illegal:/[^\s:]/,
contains:[a,E,e.C_LINE_COMMENT_MODE,e.C_BLOCK_COMMENT_MODE]},{className:"meta",
begin:"^\\s*\\[(?=[\\w])",excludeBegin:!0,end:"\\]",excludeEnd:!0,contains:[{
className:"string",begin:/"/,end:/"/}]},{
beginKeywords:"new return throw await else",relevance:0},{className:"function",
begin:"("+_+"\\s+)+"+e.IDENT_RE+"\\s*(<[^=]+>\\s*)?\\(",returnBegin:!0,
end:/\s*[{;=]/,excludeEnd:!0,keywords:n,contains:[{
beginKeywords:"public private protected static internal protected abstract async extern override unsafe virtual new sealed partial",
relevance:0},{begin:e.IDENT_RE+"\\s*(<[^=]+>\\s*)?\\(",returnBegin:!0,
contains:[e.TITLE_MODE,E],relevance:0},{match:/\(\)/},{className:"params",
begin:/\(/,end:/\)/,excludeBegin:!0,excludeEnd:!0,keywords:n,relevance:0,
contains:[g,i,e.C_BLOCK_COMMENT_MODE]
},e.C_LINE_COMMENT_MODE,e.C_BLOCK_COMMENT_MODE]},b]}}})()
;hljs.registerLanguage("csharp",e)})();
/*****************************************************************************************/

// The following is for emacs. It must be at the end of the file and is
// needed to preserve the BOM mark when editing in emacs. The begin and
// end comment on each line are also required.
 
/* Local Variables: */
/* mode: JavaScript */
/* coding: utf-8-with-signature */
/* End: */

