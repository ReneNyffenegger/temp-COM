Apartements were referred to as *execution context* in the original COM specification.

In a multithreaded environment, COM objects use Apartments to synchronize
access to resources that may be shared by multiple threads.

Managed code
objects use synchronization regions and synchronization primitives like
monitors , mutexes , and locks for the same purpose. Yes, you read that
correctly. Assuming that you are not using COM Interop (or Windows Forms), your
managed code does not use Apartments. For those of you who never quite grasped
the concept of COM Apartments, this is probably cause for celebration .

Note

I personally think that Microsoft dropping the Apartment concept in the .NET
Framework is a tacit admission by them that the "fix"COM Apartments was worse
than the original problem the need to synchronize access to resources in a
multithreaded environment.


But before you pop the champagne , keep in mind that, if you are reading this
book, you probably want to use COM components in your managed code
applications, and, unfortunately , if you do use COM components, your managed
code application will use Apartments.

Before a thread can instantiate a COM object, it must enter an Apartment. The
way I like to think of it is that entering an Apartment is just the thread's
way of announcing to the world (or at least to the COM objects that it will
subsequently create) one of two things: whether the thread is primarily an
independent thread that will not need to share resources extensively with other
threads running in the same process or it is one of a group of threads in a
multithreaded environment. An independent thread should enter an STA by calling
CoInitializeEx as follows :

 CoInitializeEx(NULL,COINIT_APARTMENT_THREADED) 
 If the thread is one thread in a multithreaded environment, it should enter the MTA (there's only one in each process) by calling CoInitializeEx as follows:

  CoInitializeEx(NULL,COINIT_MULTI_THREADED) 
  In-Process COM objects specify which type of threading environment they can operate in by adding a registry value called ThreadingModel beneath their InprocServer32 subkey . The object can specify one of the five values shown in Table 7-9.

  Note

  Out-of-process (executable) servers do not use the ThreadingModel registry key because they do not interact directly with their client's threads because the client runs in a different process.
