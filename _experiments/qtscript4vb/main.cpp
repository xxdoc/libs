#include <QApplication>
#include <QUiLoader>
#include <QtScript>
#include <QWidget>
#include <QFile>
#include <QMainWindow>
#include <QLineEdit>
#include <QAction>

#include <QScriptEngineDebugger>
#include <windows.h>
#include <qaxtypes.h>

int argc = 0;
char* argv = 0;
const short VB_TRUE = -1;
const short VB_FALSE = 0;

QApplication app(argc,&argv);
QScriptEngine *engine = new QScriptEngine();
QScriptEngineDebugger* debugger = new QScriptEngineDebugger();

int mTimeout = -1;
bool UseDebugger = false;
HANDLE hThread=0;
bool engine_init = false;
QScriptValue *qsv_retVal=0;
bool shutting_down = false;

//#pragma comment(linker, "/NODEFAULTLIB:LIBCMTD.lib") 
#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

typedef int (__stdcall *vbHostResolverCallback)(const char*, int, int, int); //*string, ctx, arg_cnt, hInst
vbHostResolverCallback vbHostResolver = 0;

enum op{
	op_reset = 0,
	op_setdbg = 1,
	op_setTimeout = 2,
	op_setResolverHandler = 3,
	op_setRetValInt = 4,
	op_setRetValStr = 5,
	op_getVarFromCtx = 6,
	op_qtShutdown = 7
};

QScriptValue comResolver(QScriptContext *context , QScriptEngine *engine)
{
	QString arg0;
	std::string s_arg0;

	if(vbHostResolver==0){
		MessageBoxA(0,"vbHostResolver not yet set!","",0);
		goto retNull;
	}

	int realArgCount = context->argumentCount();
	if(realArgCount==0) goto retNull;
	
	//this has to be on 3 seperate lines or didnt work..
	arg0 = context->argument(0).toString();
	s_arg0 = arg0.toStdString(); 
	const char* meth = s_arg0.c_str();

	int hasRetVal = vbHostResolver(meth, (int)context, realArgCount, 0);
	if(hasRetVal==0 || (int)qsv_retVal==0) goto retNull;

	return *qsv_retVal;

retNull:
	return QScriptValue::NullValue;
}

QScriptValue myAlert(QScriptContext *context , QScriptEngine *engine)
{
	MessageBoxA(0, context->argument(0).toString().toAscii(),"QtScript Alert",0);
	return QScriptValue::NullValue;
}

QScriptValue nativeToUpper(QScriptContext *context , QScriptEngine *engine)
{
	QString tmp = QString("in nativeToUpper argCount= %1 (").arg(context->argumentCount());
    
	// - or -
	//QString tmp = QLatin1String("in nativeToUpper ");
	//QTextStream(&tmp) << "argCount=" << context->argumentCount() << " (";
	
	for(int i=0; i< context->argumentCount(); i++){
		tmp += context->argument(i).toString()+", ";
	}

	tmp.chop(2);
	tmp.append(")");
	MessageBoxA(0,tmp.toAscii(),"",0);

	QScriptValue s(context->argument(0).toString().toUpper() );
    return s;
}

void registerNative(bool force = false)
{
	if(engine_init && !force) return;

	QScriptValue globalObject = engine->globalObject();
	
	QScriptValue func = engine->newFunction(nativeToUpper);
	globalObject.setProperty("nativeToUpper", func);

	QScriptValue func2 = engine->newFunction(myAlert);
	globalObject.setProperty("alert", func2);
	
	QScriptValue func3 = engine->newFunction(comResolver);
	globalObject.setProperty("resolver", func3);
	
	QMainWindow *debugWindow = debugger->standardWindow();
	debugWindow->resize(1024, 640);

	engine_init = true;

}

void BreakNext(){//break on next instruction
	QAction *qa = debugger->action(QScriptEngineDebugger::InterruptAction);
	qa->trigger(); 
}

unsigned int __stdcall QtOp(int operation,unsigned int v1, unsigned int v2, unsigned int v3)
{
#pragma EXPORT

	QScriptContext *context = 0;
	QScriptValue v;
	VARIANT *vv=0;

	if(shutting_down) return VB_FALSE;
	//if(engine->isEvaluating()) return VB_FALSE; some operations are designed to be used at runtime..no top level check!

	switch(operation){
		case op_setResolverHandler:

				vbHostResolver = (vbHostResolverCallback )v1;
				break;

		case op_reset: 
				
				debugger->detach();
				delete debugger;
				delete engine;
				engine = new QScriptEngine();
				debugger = new QScriptEngineDebugger();
				registerNative(true); 
				if(UseDebugger){
					debugger->attachTo(engine);
					BreakNext();
				}
				break; 

		case op_setdbg : 

				UseDebugger = (v1==1) ? true : false; 

				if(UseDebugger){
					debugger->attachTo(engine);
					BreakNext();
				}else{
					debugger->detach();
				}
				
				break;

		case op_setTimeout: 
				
				mTimeout = v1; 
				break;

		case op_setRetValInt: 
				
				if((int)qsv_retVal != 0) delete qsv_retVal;
				qsv_retVal = new QScriptValue(v1);
				break;

		case op_setRetValStr: 
				
				if((int)qsv_retVal != 0) delete qsv_retVal;
				qsv_retVal = new QScriptValue((char*)v1);
				break;

		case op_getVarFromCtx:
				
				if(v1==0 || v3==0) return VB_FALSE;
				context = (QScriptContext*)v1;
				if(v2 > context->argumentCount()) return VB_FALSE;
				v = context->argument(v2);
				vv = (VARIANT*)v3;
				QVariantToVARIANT(v.toVariant(),*vv);
				break;

		case op_qtShutdown:
				
				shutting_down = true;
				if((int)qsv_retVal != 0) delete qsv_retVal;
				if(hThread!=0) TerminateThread(hThread,0);
				vbHostResolver = (vbHostResolverCallback )0;
				debugger->detach();
				if(engine->isEvaluating()) engine->abortEvaluation();
				delete debugger;
				delete engine;
				app.exit();

			break;
			/*
		case op_and: return v1 & v2;
		case op_or:  return v1 | v2;*/
	}

	return VB_TRUE;

}


DWORD WINAPI MyScriptTimeOut( LPVOID lpParam ) 
{ 
	DWORD start = GetTickCount();
	char *msg = "Script engine timeout has expired.\nContinue waiting or abort?";

try_again:
	while( (GetTickCount() - start) < (int)lpParam){if(shutting_down) return VB_FALSE;}
	
	if( MessageBoxA(0,msg,"Script Timeout", MB_RETRYCANCEL) == IDRETRY){
		start = GetTickCount();
		goto try_again;
	}

	if(engine->isEvaluating()) engine->abortEvaluation(0);
	//if the user used ::call from C, above will not work have to throw error
	return 1;
}

short __stdcall AddFile(char* fPath)
{
#pragma EXPORT

	if(shutting_down) return VB_FALSE;
	if(engine->isEvaluating()) return VB_FALSE;

	registerNative();
	QString scriptFileName(fPath);
    QFile scriptFile(scriptFileName);
	if(!scriptFile.exists()) return VB_FALSE;

    scriptFile.open(QIODevice::ReadOnly);
	QString code = scriptFile.readAll();
	scriptFile.close();

	if(UseDebugger) BreakNext();
	if(mTimeout > 0 && !UseDebugger) HANDLE hThread = CreateThread(NULL, 0, MyScriptTimeOut, (LPVOID)4000, 0, 0);  
    QScriptValue v = engine->evaluate(code, scriptFileName);
	if(hThread!=0){ TerminateThread(hThread,0); hThread=0;}
	if(UseDebugger) debugger->standardWindow()->hide();

	if (engine->hasUncaughtException()) {
		int line = engine->uncaughtExceptionLineNumber();
		QString tmp = QString("uncaught exception at line %1\n").arg(line);
		tmp += "Description: " + v.toString()+"\n";
		tmp += "BackTrace: " + engine->currentContext()->backtrace().join("\n");
		MessageBoxA(0,tmp.toAscii(),"",0);
		return VB_FALSE;
	}

	return VB_TRUE;
}

short __stdcall Eval(char* vbstr, VARIANT *vvv)
{
#pragma EXPORT
	
	if(shutting_down) return VB_FALSE;
	if(engine->isEvaluating()) return VB_FALSE;

	registerNative();
	VariantClear(vvv);	
    QString code = QLatin1String(vbstr) ;

	if(UseDebugger) BreakNext();
	if(mTimeout > 0 && !UseDebugger) HANDLE hThread = CreateThread(NULL, 0, MyScriptTimeOut, (LPVOID)4000, 0, 0);    
    QScriptValue v = engine->evaluate(code, "dummy");
	if(hThread!=0){ TerminateThread(hThread,0); hThread=0;}
	if(UseDebugger) debugger->standardWindow()->hide();

	if (engine->hasUncaughtException()) {
		int line = engine->uncaughtExceptionLineNumber();
		QString tmp = QString("uncaught exception at line %1\n").arg(line);
		tmp += "Description: " + v.toString()+"\n";
		tmp += "BackTrace: " + engine->currentContext()->backtrace().join("\n");
		MessageBoxA(0,tmp.toAscii(),"",0);
		return VB_FALSE;
	}

	QVariantToVARIANT(v.toVariant(),*vvv);

	return VB_TRUE;

}