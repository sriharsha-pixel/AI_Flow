const { log } = require("console");

async function renameFile(page,element,rename){
    const fileName=await element.inputValue();
    const parts=fileName.split('.');
    const extension=parts.pop();
    const newfileName=`${rename}.${extension}`;
    console.log("file is",fileName);
    console.log("extension is ",extension);
    console.log("new file name is ",newfileName);
    return newfileName;
}

async function invalidrenameFile(page,element,rename){
    const fileName=await element.inputValue();
    const parts=fileName.split('.');
    const extension=parts.pop();
    const newExtension=extension.toUpperCase();
    const newfileName=`${rename}.${newExtension}`;
    return newfileName;
}
module.exports={renameFile,invalidrenameFile};