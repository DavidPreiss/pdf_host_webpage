let firstname = "firstname";
let lastname = "lastname";

function dispName()
{
    firstname = document.getElementById("userInput").value;
	const newURL = "./User"+firstname+".html";
    document.getElementById("nameshower").innerHTML=newURL;
	
	const mylink = document.getElementById("namelink");
	mylink.href = newURL
	
	const mylink2 = document.getElementById("embedded_pdf");
	mylink2.href = newURL
}
