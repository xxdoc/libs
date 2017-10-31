import zmq
import sys

port = "5546"
if len(sys.argv) > 1:
    port =  sys.argv[1]
    int(port)

if len(sys.argv) > 2:
    port1 =  sys.argv[2]
    int(port1)

context = zmq.Context()
print ("Connecting to server...on port %s" % port)
socket = context.socket(zmq.REQ)
socket.connect ("tcp://localhost:%s" % port)

if len(sys.argv) > 2:
    socket.connect ("tcp://localhost:%s" % port1)

print "Sending request " 
socket.setsockopt(zmq.LINGER, 0)

# use poll for timeouts:
poller = zmq.Poller()
poller.register(socket, zmq.POLLIN)
socket.send ("Hello")
    
if poller.poll(3*1000): #  timeout in milliseconds
    message = socket.recv()
else:
    raise IOError("Timeout processing auth request")

#  Get the reply.

print "Received reply ",  "[", message, "]"

print "done"