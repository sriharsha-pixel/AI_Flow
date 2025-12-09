exports.excuteSteps = async (test, element, action, discription, data) => {
  await test.step(discription, async () => {
    switch (action) {
      case "click":
        await element.click();
        break;
      case "fill":
        await element.fill(data[0]);
        break;
      case "dblclick":
        await element.dblclick();
        break;
      case "navigate":
        await element.goto(data[0]);
        break;
      case "type":
        await element.type(data[0]);
        break;
      case "check":
        await element.check();
        break;
      case "tap":
        await element.tap();
        break;
      case "hover":
        await element.hover();
        break;
      case "scroll":
        await element.scrollIntoViewIfNeeded();
        break;
      
    }
  });
};


