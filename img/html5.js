!function(e,n){var t=function(e){return e.innerHTML="<x-element></x-element>",1===e.childNodes.length}(n.createElement("a")),i=function(e,n,t){return n.appendChild(e),(t=(t?t(e):e.currentStyle).display)&&n.removeChild(e)&&"block"===t}(n.createElement("nav"),n.documentElement,e.getComputedStyle),a={elements:"abbr article aside audio bdi canvas data datalist details figcaption figure footer header hgroup mark meter nav output progress section summary time video".split(" "),shivDocument:function(e){if(e=e||n,!e.documentShived){e.documentShived=!0;var o=e.createElement,r=e.createDocumentFragment,l=e.getElementsByTagName("head")[0];if(t||(a.elements.join(" ").replace(/\w+/g,function(e){o(e)}),e.createElement=function(e){return e=o(e),e.canHaveChildren&&a.shivDocument(e.document),e},e.createDocumentFragment=function(){return a.shivDocument(r())}),!i&&l){var c=o("div");c.innerHTML="x<style>article,aside,details,figcaption,figure,footer,header,hgroup,nav,section{display:block}audio{display:none}canvas,video{display:inline-block;*display:inline;*zoom:1}[hidden]{display:none}audio[controls]{display:inline-block;*display:inline;*zoom:1}mark{background:#FF0;color:#000}</style>",l.insertBefore(c.lastChild,l.firstChild)}return e}}};a.shivDocument(n),e.html5=a}(this,document);