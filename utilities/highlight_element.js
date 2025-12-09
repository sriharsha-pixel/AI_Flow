exports.highlightElement = async (page, locator) => {
    await page.evaluate((element) => {
      element.style.border = '2px solid red'; // Add a red border for visual highlighting
     // element.style.backgroundColor = 'yellow';
    }, await locator.elementHandle());
  };
  exports.highlighterRemove = async (page,locator) =>{
    await page.evaluate((element) => {
      element.style.border = '';
      element.style.backgroundColor = '';
    }, await locator.elementHandle())
  }
 