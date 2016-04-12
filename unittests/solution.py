#!python
# ---------------------------------------------------------------------------------------------
#  Python / Skype4Py example that takes a skypename from command line parameter,
#  checks if that skypename is in contact list and if yes then starts a call to that skypename.
#
#  Tested with  Skype4Py version 0.9.28.2 and Skype verson 3.5.0.214

import sys
import Skype4Py

# This variable will get its actual value in OnCall handler
CallStatus = 0
logfile= open("Logfile.txt", 'w');

# Here we define a set of call statuses that indicate a call has been either aborted or finished
CallIsFinished = set ([Skype4Py.clsFailed, Skype4Py.clsFinished, Skype4Py.clsMissed, Skype4Py.clsRefused, Skype4Py.clsBusy, Skype4Py.clsCancelled]);

def AttachmentStatusText(status):
   return skype.Convert.AttachmentStatusToText(status)

def CallStatusText(status):
    return skype.Convert.CallStatusToText(status)

# This handler is fired when status of Call object has changed
def OnCall(call, status):
    global CallStatus
    CallStatus = status
    print 'Call status: ' + CallStatusText(status)

# This handler is fired when Skype attatchment status changes
def OnAttach(status): 
    print 'API attachment status: ' + AttachmentStatusText(status)
    if status == Skype4Py.apiAttachAvailable:
        skype.Attach()
       
def monitor():
	pass


# Let's see if we were started with a command line parameter..
try:
    CmdLine = sys.argv[1]
except:
    print 'Missing command line parameter'
    sys.exit()

# Creating Skype object and assigning event handlers..
skype = Skype4Py.Skype()
skype.OnAttachmentStatus = OnAttach
skype.OnCallStatus = OnCall

# Starting Skype if it's not running already..
if not skype.Client.IsRunning:
    print 'Starting Skype..'
    skype.Client.Start()

# Attatching to Skype..
print 'Connecting to Skype..'
skype.Attach()
		
firstTry = True
dropcount=0;
successcount=0;
while (firstTry or CallStatus == Skype4Py.clsCancelled):
	firstTry = False
	# Checking if what we got from command line parameter is present in our contact list
	Found = False
	for F in skype.Friends:
		if F.Handle == CmdLine:
			Found = True
			print 'Calling ' + F.Handle + '..'
			skype.PlaceCall(CmdLine)
			print 'Call placed'
			break

	if not Found:
		print 'Call target not found in contact list'
		sys.exit()

	# Loop until CallStatus gets one of "call terminated" values in OnCall handler
	#while not CallStatus in CallIsFinished:
	#	print CallStatus
	#	pass
	while CallStatus!=Skype4Py.clsCancelled:
		# print CallStatus	
		if CallStatus=='INPROGRESS':
			monitor();
		if CallStatus=='FINISHED':
			successcount+=1;
			#print dropcount;
			logfile.write("Number of dropcounts: " +str(dropcount));
			logfile.write("\n");
			logfile.write("Number of successful calls: " +str(successcount));
			logfile.close();
			sys.exit()
		pass
	dropcount+=1;
	try:
		while len(skype.ActiveCalls)>0:
			skype.ActiveCalls[0].Finish()
	except:
		print 'No active calls in queue'