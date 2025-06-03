function dispName()
{
    const firstname = document.getElementById("userInput").value.toUpperCase();
	//const newURL = "./Users/User"+firstname+"/User"+firstname+".html";
	const newURL = "./Customers/"+firstname+"/"+firstname+".html";
    document.getElementById("nameshower").innerHTML=newURL;
	
	const mylink = document.getElementById("namelink");
	mylink.href = newURL
	
}
