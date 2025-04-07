const {activeXObject}= require('winax')
try {
    const outLook= new activeXObject('Outlook.Application');
    const mapi=outLook.getNamespace('MAPI');
    const inbox=mapi.GetDefaultFolder(6);
    const items=inbox.Items;
    if(items.Count>0){
        const firstMail=items.Item(1);
        console.log(firstMail.Subject);
    }
    else{
        console.log("Empty");
        
    }
} catch (error) {
    console.log("Bye");
       
}