async function scrollToElement(locator, position = "center") {
  await locator.waitFor({ state: "attached" }); // ensure element is in DOM

  await locator.evaluate((el, pos) => {
    el.scrollIntoView({
      block: pos,     // "center", "start", "end", "nearest"
      inline: pos,    // same for horizontal scrolling
      behavior: "instant" // no smooth scrolling in automation
    });
  }, position);
}

module.exports = { scrollToElement };
