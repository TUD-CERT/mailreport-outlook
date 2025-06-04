/* global document, Element, Node, Office */
/*
 * Derived from:
 * http://github.com/piroor/webextensions-lib-l10n
 *
 * Original license:
 * The MIT License, Copyright (c) 2016-2019 YUKI "Piro" Hiroshi
 */
import locales from "./locales.json";

const keyPrefix = "__MSG_";

export function localizeToken(token: string) {
  let re = new RegExp(keyPrefix + "(.+?)__", "g");
  return token.replace(re, (matched) => {
    const key = matched.slice(keyPrefix.length, -2);
    let language = Office.context.displayLanguage.split("-")[0];
    if (!(language in locales)) language = "en";
    return locales[language][key];
  });
}

function updateSubtree(node: Node) {
  if (node.nodeType === Node.TEXT_NODE) {
    if (node.nodeValue && node.nodeValue.includes(keyPrefix)) node.nodeValue = localizeToken(node.nodeValue);
  } else {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const element = node as Element;
      for (let i = 0; i < element.attributes.length; i++) {
        const attr = element.attributes[i];
        if (attr.value.includes(keyPrefix)) attr.value = localizeToken(attr.value);
      }
    }
    for (let i = 0; i < node.childNodes.length; i++) {
      updateSubtree(node.childNodes[i]);
    }
  }
}

export function localizeDocument() {
  updateSubtree(document);
  document.body.classList.remove("hide");
}
