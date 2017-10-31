
zmq is an interprocess communications library that people seem to like

this is a test project for using libzmq between python and vb6
the python was taken from the web and you have to do a pip install

the libzmq.dll is build as stdcall for vb6 and includes the polling draft api
and was built to also support down to xp

cant seem to get the polling api to work right now it still hangs bad if
the server.py isnt running

it does work if the server is up. if you lose your connection right now your
fucked though...

needs an easier timeout mechanism for sync requests..
i am switching to something else but maybe this can help you 


http://learning-0mq-with-pyzmq.readthedocs.io/en/latest/pyzmq/patterns/client_server.html

http://zeromq.org/bindings:python

# (Windows or OS X)
pip install --wheel pyzmq
# or
easy_install pyzmq
# or (pretty much anywhere)
pip install pyzmq

dl