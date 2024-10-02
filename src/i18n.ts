/* global Attr, document, Node, Office, XPathResult */
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
    const language = Office.context.displayLanguage;
    return locales[language][key];
  });
}

function updateSubtree(node: Node) {
  const texts = document.evaluate(
    'descendant::text()[contains(self::text(), "' + keyPrefix + '")]',
    node,
    null,
    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,
    null
  );
  for (let i = 0; i < texts.snapshotLength; i++) {
    const text = texts.snapshotItem(i);
    if (text.nodeValue.includes(keyPrefix)) text.nodeValue = localizeToken(text.nodeValue);
  }

  const attributes = document.evaluate(
    'descendant::*/attribute::*[contains(., "' + keyPrefix + '")]',
    node,
    null,
    XPathResult.ORDERED_NODE_SNAPSHOT_TYPE,
    null
  );
  for (let i = 0; i < attributes.snapshotLength; i++) {
    const attribute = <Attr>attributes.snapshotItem(i);
    if (attribute.value.includes(keyPrefix)) attribute.value = localizeToken(attribute.value);
  }
}

export function localizeDocument() {
  updateSubtree(document);
  document.body.classList.remove("hide");
}
