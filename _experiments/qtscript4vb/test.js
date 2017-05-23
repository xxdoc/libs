
function a(x){
	return x+1;
}

function b(x){
	return a(x+1)
}

v = b(1);
alert(v);