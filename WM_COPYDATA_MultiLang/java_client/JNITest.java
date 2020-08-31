
@SuppressWarnings("unchecked") 

public class JNITest{

	public native int InitHwnd();
	public native String SendCMDRecvText(String msg);
	public native int SendCMDRecvInt(String msg);
	public native void SendCMD(String msg); 
	
	static { System.loadLibrary("copydata"); }   
	
	private int hwnd=0;

	public void AsyncMsg(String msg){ //test for async callback 
		System.out.println("JAVA: AsyncMessage("+ msg +")");
	}

	public static void main(String[] args) 
	{	 
		JNITest t = new JNITest();
		System.out.println("JAVA: Trying to find VB6 target window..");
		
		t.hwnd = t.InitHwnd(); 
		
		if(t.hwnd==0){
			System.out.println("Init failed to create command window and/or find vb6 target");
			return;
		}
		
		System.out.println("JAVA: sending hello message");
		t.SendCMD("Hello from java!");
		
		String response = t.SendCMDRecvText("PINGME="+t.hwnd);
		System.out.println("JAVA: Response from VB: " + response);
		  
		System.out.println("JAVA: Going into wait loop for async messages..");
		while(true){ try{t.wait();}catch(Exception e){} }

		
	}

}
