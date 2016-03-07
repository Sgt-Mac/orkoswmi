//////////////////////////////////////////////////////////////////////////////-
//- H E A D E R S
//////////////////////////////////////////////////////////////////////////////-
#define _WIN32_DCOM
#include <string>
#include <comdef.h>
#include <Wbemidl.h>

#include <stdio.h>
#include <python.h>
#include <structmember.h>

#pragma comment( lib, "wbemuuid.lib")
#pragma comment(lib, "mpr.lib")


using namespace std;

//////////////////////////////////////////////////////////////////////////////-
//- D E F I N E S
//////////////////////////////////////////////////////////////////////////////-
//////////////////////////////////////////////////////////////////////////////-
//- T Y P E S
//////////////////////////////////////////////////////////////////////////////-
//////////////////////////////////////////////////////////////////////////////-
//- PyWMI Type Object
typedef struct PYWMIOBJECT
{
  PyObject_HEAD

  PyObject *host;

  IWbemServices *cimv2_svc  ;
  IWbemServices *default_svc;

  IWbemClassObject *win32_process_class;
  IWbemClassObject *stdregprov_class   ;

} PyWMIObject;


//////////////////////////////////////////////////////////////////////////////-
//- S T A T I C  F U N C T I O N  P R O T O T Y P E S
//////////////////////////////////////////////////////////////////////////////-
PyObject *PyWMI_new        ( PyTypeObject *type, PyObject *args, PyObject *kwargs );
void      PyWMI_dealloc    ( PyWMIObject *self);
int       PyWMI_init       ( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_connect    ( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_file_create( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_file_write ( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_file_close ( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_file_exec  ( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_disconnect ( PyWMIObject *self);
PyObject *PyWMI_file_delete( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_file_get   ( PyWMIObject *self , PyObject *args, PyObject *kwargs );
PyObject *PyWMI_reg_getexpandedstringvalue(PyWMIObject *self,PyObject *args, PyObject *kwargs );

//////////////////////////////////////////////////////////////////////////////-
//- S T A T I C  V A R I A B L E S
//////////////////////////////////////////////////////////////////////////////-
//////////////////////////////////////////////////////////////////////////////-
static char PyWMI_doc[] = "We should really write some documentation.";
//////////////////////////////////////////////////////////////////////////////-
//- PyWMI object members
static PyMemberDef PyWMI_members[] =
{
  { "host", T_OBJECT_EX, offsetof( PyWMIObject, host    ), 0, "host"},
  { NULL}
};

//////////////////////////////////////////////////////////////////////////////-
//- PyWMI object methods
#define _c(x) ((PyCFunction)x)
static PyMethodDef PyWMI_methods[] =
{
{"connect"    , _c(PyWMI_connect    ), METH_VARARGS | METH_KEYWORDS,"Connect to host"        },
{"file_create", _c(PyWMI_file_create), METH_VARARGS | METH_KEYWORDS,"Creates file on host"   },
{"file_write" , _c(PyWMI_file_write ), METH_VARARGS | METH_KEYWORDS,"Writes to file "        },
{"file_close" , _c(PyWMI_file_close ), METH_VARARGS | METH_KEYWORDS,"Closes file..."         },
{"file_exec"  , _c(PyWMI_file_exec  ), METH_VARARGS | METH_KEYWORDS ,"Executes file on host" },
{"disconnect" , _c(PyWMI_disconnect ), METH_NOARGS                  ,"Disconnects from host" },
{"file_delete", _c(PyWMI_file_delete), METH_VARARGS | METH_KEYWORDS ,"Deletes file on host"  },
{"file_get"   , _c(PyWMI_file_get   ), METH_VARARGS | METH_KEYWORDS ,"Copies file from host" },
{"reg_GetExpandedStringValue", _c(PyWMI_reg_getexpandedstringvalue), METH_VARARGS | METH_KEYWORDS ,"Returns data value from remote registry" },
{ NULL}
};


//////////////////////////////////////////////////////////////////////////////-
//- PyWMI Type Object
static PyTypeObject PyWMIObjectType =
{
  PyObject_HEAD_INIT(NULL)
  0,                              /* ob_size           */
  "orkoswmi.PyWMI",               /* tp_name           */
  sizeof(PyWMIObject),            /* tp_basicsize      */
  0,                              /* tp_itemsize       */
  (destructor)PyWMI_dealloc,      /* tp_dealloc        */
  0,                              /* tp_print          */
  0,                              /* tp_getattr        */
  0,                              /* tp_setattr        */
  0,                              /* tp_compare        */
  0,                              /* tp_repr           */
  0,                              /* tp_as_number      */
  0,                              /* tp_as_sequence    */
  0,                              /* tp_as_mapping     */
  0,                              /* tp_hash           */
  0,                              /* tp_call           */
  0,                              /* tp_str            */
  0,                              /* tp_getattro       */
  0,                              /* tp_setattro       */
  0,                              /* tp_as_buffer      */
  Py_TPFLAGS_DEFAULT | Py_TPFLAGS_BASETYPE, /* tp_flags          */
  PyWMI_doc,                      /* tp_doc            */
  0,                              /* tp_traverse       */
  0,                              /* tp_clear          */
  0,                              /* tp_richcompare    */
  0,                              /* tp_weaklistoffset */
  0,                              /* tp_iter           */
  0,                              /* tp_iternext       */
  PyWMI_methods,                  /* tp_methods        */
  PyWMI_members,                  /* tp_members        */
  0,                              /* tp_getset         */
  0,                              /* tp_base           */
  0,                              /* tp_dict           */
  0,                              /* tp_descr_get      */
  0,                              /* tp_descr_set      */
  0,                              /* tp_dictoffset     */
  (initproc)PyWMI_init,           /* tp_init           */
  0,                              /* tp_alloc          */
  PyWMI_new                       /* tp_new            */
};


//////////////////////////////////////////////////////////////////////////////-
//- Function
//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Helper function/wrapper for CoInitialize/CoInitializeSecurity
//-
//- Arguments:
//-   void
//-
//- Return Value:
//-   int
//-    - Failure: 0
//-    - Success: 1
//////////////////////////////////////////////////////////////////////////////-
static unsigned int com_ref_count = 0;
static int
_init_com()
{
  int retval = 0;

  //- Init COM
  if( com_ref_count <= 0)
  {
    HRESULT hres = CoInitializeEx( 0, COINIT_MULTITHREADED);
    if( FAILED( hres))
    {
        PyErr_Format( PyExc_RuntimeError, "CoInitializeEx failed with Error Code: [0x%x](%ld)", hres, hres);
      goto _INIT_COM_ERROR;
    }
    //- Init COM security levels
    hres = CoInitializeSecurity(
                                 NULL,
                                 -1,                          // COM authentication
                                 NULL,                        // Authentication services
                                 NULL,                        // Reserved
                                 RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
                                 RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
                                 NULL,                        // Authentication info
                                 EOAC_NONE,                   // Additional capabilities 
                                 NULL                         // Reserved
    );
    if( FAILED( hres))
    {
      PyErr_Format( PyExc_RuntimeError, "CoInitializeSecurity failed with Error Code: [0x%x](%ld)", hres, hres);
      goto _INIT_COM_ERROR;
    }
    retval = 1;
  }

_INIT_COM_END:
  com_ref_count++;
  return( retval);

_INIT_COM_ERROR:
  goto _INIT_COM_END;

}


//////////////////////////////////////////////////////////////////////////////-
//- Function
//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Helper function/wrapper for CoUninitialize
//-
//- Arguments:
//-   void
//-
//- Return Value:
//-   void
//////////////////////////////////////////////////////////////////////////////-
static void
_uninit_com()
{
  if( com_ref_count <= 0)
  {
    CoUninitialize();
  }
  com_ref_count--;
}


//////////////////////////////////////////////////////////////////////////////-
//////////////////////////////////////////////////////////////////////////////-
static PyMethodDef module_methods[] =
{
  {NULL}  /* Sentinel */
};
//////////////////////////////////////////////////////////////////////////////-
//- Function
//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Initializes orkoswmi module, called auto-magically by Python
//-
//- Arguments:
//-   void
//-
//- Return Value:
//-   void
//////////////////////////////////////////////////////////////////////////////-
PyMODINIT_FUNC
initorkoswmi( void)
{
  PyObject *module = 0;
  PyObject *dict   = 0;

  if( PyType_Ready( &PyWMIObjectType) < 0)
    return;

  module = Py_InitModule3( "orkoswmi", module_methods, "Module that creates orkoswmi PyWMI objects");
  if( module == NULL)
  {
    goto INITORKOSWMI_ERROR;
  }

  Py_INCREF( &PyWMIObjectType);
  dict = PyModule_GetDict( module);

  PyModule_AddObject( module, "PyWMI", (PyObject *)&PyWMIObjectType);

  //- Adding module globals
  //- file create access
  PyDict_SetItemString( dict, "GENERIC_READ"   , PyLong_FromLong( GENERIC_READ   ));
  PyDict_SetItemString( dict, "GENERIC_WRITE"  , PyLong_FromLong( GENERIC_WRITE  ));
  PyDict_SetItemString( dict, "GENERIC_EXECUTE", PyLong_FromLong( GENERIC_EXECUTE));
  PyDict_SetItemString( dict, "GENERIC_ALL"    , PyLong_FromLong( GENERIC_ALL    ));

  //- file create disposition
  PyDict_SetItemString( dict, "CREATE_NEW"       , PyLong_FromLong( CREATE_NEW   ));
  PyDict_SetItemString( dict, "CREATE_ALWAYS"    , PyLong_FromLong( CREATE_ALWAYS));
  PyDict_SetItemString( dict, "OPEN_EXISTING"    , PyLong_FromLong( OPEN_EXISTING));
  PyDict_SetItemString( dict, "OPEN_ALWAYS"      , PyLong_FromLong( OPEN_ALWAYS  ));
  PyDict_SetItemString( dict, "TRUNCATE_EXISTING", PyLong_FromLong( TRUNCATE_EXISTING));

INITORKOSWMI_END:
  //return( module);
  return;

INITORKOSWMI_ERROR:
  Py_DECREF( module);
  module = 0;
  goto INITORKOSWMI_END;
}


//////////////////////////////////////////////////////////////////////////////-
//- Function
//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Allocates memory for PyWMI object
//-
//- Arguments:
//-   void
//-
//- Return Value:
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: PyWMIObject fully initialized ready for use via Python
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_new( PyTypeObject *type, PyObject *args, PyObject *kwargs )
{
  PyWMIObject *self = 0;

  if( type)
  {
    self = (PyWMIObject *)(type->tp_alloc( type, 0));
    if( self)
    {
      self->host = PyString_FromString( "");
      if( !(self->host))
      {
        goto PYWMI_NEW_ERROR;
      }
    }
  }
  _init_com();

PYWMI_NEW_END:
  return( (PyObject *)self);

PYWMI_NEW_ERROR:

  if( self)
  {
    if( (self->host))
    {
      Py_DECREF( self->host);
    }
    Py_DECREF( self);
    self = 0;
  }
  goto PYWMI_NEW_END;

}


//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Deallocates PyWMI object
//-
//- Arguments:
//-   void
//-
//- Return Value:
//-  void
//////////////////////////////////////////////////////////////////////////////-
static void
PyWMI_dealloc( PyWMIObject *self)
{
  if( self)
  {
    Py_XDECREF(self->host    );
    self->ob_type->tp_free((PyObject*)self);
  }
  _uninit_com();
}


//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Initializes the PyWMI object
//-
//- Arguments:
//-  - PyObject (optional): host: the remote host address or hostname
//-  - - Note: A host *MUST* be defined either by init or connect
//-
//- Return Value:
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static int
PyWMI_init( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  int retval = 0;

  char *kwlist[] = { "host", NULL };

  if( self)
  {
    PyObject *host     = 0;

    retval = PyArg_ParseTupleAndKeywords( args, kwargs, "|S", kwlist, &(host));
    if( !retval)
    {
      goto PYWMI_INIT_ERROR;
    }
    if( self->host)
    {
      Py_DECREF( self->host);
    }
    if( host)
    {
      Py_INCREF( host);
      self->host = host;
    }
    retval = 1;
  }

PYWMI_INIT_END:
  return( retval);

PYWMI_INIT_ERROR:
  retval = 0;
  goto PYWMI_INIT_END;
}


//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Connects to remote wmi services
//-
//- Arguments:
//-  - PyObject (optional): host    : the remote host address or hostname
//-  - PyObject (optional): username: the username to login to remote host
//-  - PyObject (optional): password: the usernames password
//-  - PyObject (optional): do_kerb : Use kerberos (defaults to 'False')
//-
//- Return Value:
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_connect( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;

  //- upn == User Principal Name (used for Kerberos Authenticion only)
  char *kwlist[] = { "host", "username", "password", "do_kerb", NULL };

  PyObject      *py_host      = 0;
  PyObject      *py_username  = 0;
  PyObject      *py_password  = 0;
  unsigned char  do_kerb      = 0;
  int str_size = 0;
  const wchar_t *username_str = 0;
  const wchar_t *password_str = 0;
  const wchar_t *upn_str      = 0;

  wchar_t namespac_buf[MAX_PATH+1] = {0};
  wchar_t username_buf[MAX_PATH+1] = {0};
  wchar_t password_buf[MAX_PATH+1] = {0};
  wchar_t upn_buf     [MAX_PATH+1] = {0};

  IWbemLocator  *pLoc = 0;

  if( !PyArg_ParseTupleAndKeywords( args, kwargs, "|SSSb", kwlist, &(py_host    ),
                                                                   &(py_username),
                                                                   &(py_password),
                                                                   &(do_kerb    )))
  {
    goto PYWMI_CONNECT_ERROR;
  }

  if( self->host)
  {
    Py_DECREF( self->host);
    self->host = 0;
  }
  if( py_host)
  {
    Py_INCREF( py_host);
    self->host = py_host;
  }

  //- Must have a host defined in order to be able to actually connect to a host
  if( !self->host)
  {
    PyErr_Format( PyExc_TypeError, "Remote 'host' is not defined...");
    goto PYWMI_CONNECT_ERROR;
  }

  if(PyString_Size( self->host) == 0)
  {
    PyErr_Format( PyExc_ValueError, "for argument 'host' (pos 1)");
    goto PYWMI_CONNECT_ERROR;
  }

  if( py_username)
  {
    swprintf_s( username_buf, MAX_PATH, L"%S", PyString_AsString( py_username));
    username_str = username_buf;
  }
  if( py_password)
  {
    swprintf_s( password_buf, MAX_PATH, L"%S", PyString_AsString( py_password));
    password_str = password_buf;
  }
  if( do_kerb == 1)
  {
    swprintf_s( upn_buf, MAX_PATH, L"Kerberos:%S", PyString_AsString( self->host));
    upn_str = upn_buf;
  }

  //-rjm TESTING XXX
  fprintf( stdout, "[%s:%d] username: [%s]\n", __FILE__, __LINE__, username_str);
  fprintf( stdout, "[%s:%d] password: [%s]\n", __FILE__, __LINE__, password_str);

  //- if self->cimv2_svc or self->default_svc are not NULL,
  //-  we have already made a connection therefore we must
  //-  disconnect before we can connect again
  if( self->cimv2_svc || self->default_svc )
  {
    PyWMI_disconnect( self);
  }

  //- Obtain the initial locator for WMI
  HRESULT hres = CoCreateInstance( CLSID_WbemLocator,
                                   0,
                                   CLSCTX_INPROC_SERVER,
                                   IID_IWbemLocator, (LPVOID *)&(pLoc));
  if( FAILED( hres))
  {
   PyErr_Format( PyExc_RuntimeError, "Failed to create IWbemLocator object. Error Code: [0x%x](%ld)", (LONG )hres, (LONG )hres);
    goto PYWMI_CONNECT_ERROR;
  }

  //- Connect to the root\cimv2 namespace with the current user
  //-  and obtain pointer cimv2_svc to make IWbemServices calls
  swprintf_s( namespac_buf, MAX_PATH, L"\\\\%S\\ROOT\\CIMV2", PyString_AsString( self->host));
  fprintf( stdout, "namespace_buf: [%S]\n", namespac_buf);
  hres = pLoc->ConnectServer(
        _bstr_t( namespac_buf),      // Object path of wmi namespace
        _bstr_t( username_str),      // User name. NULL = current
        _bstr_t( password_str),      // User password. NULL = current
        0   ,                        // Local.  NULL = current
        NULL,                        // Security flags.
        _bstr_t( upn_str     ),      // Authority (for example, Kerberos) RJM-FixMe?
        0   ,                        // Contect object
        &(self->cimv2_svc)           // pointer to IWbemServices proxy
    );
  if( FAILED( hres))
  {

    PyErr_Format( PyExc_RuntimeError, "Failed to create cimv2_svc IWbemLocator object for namespace: [%s]. Error Code: [0x%x](%ld)", PyString_AsString( self->host), (LONG )hres, (LONG )hres);
    goto PYWMI_CONNECT_ERROR;
  }
  pLoc->Release();

  //- Connect to the root\cimv2 namespace with the current user
  //-  and obtain pointer cimv2_svc to make IWbemServices calls
  memset( namespac_buf, 0, MAX_PATH);
  swprintf_s( namespac_buf, MAX_PATH, L"\\\\%s\\root\\default",PyString_AsString( self->host));
  fprintf( stdout, "namespace_buf: [%S]\n", namespac_buf);
  hres = pLoc->ConnectServer(
        _bstr_t( namespac_buf),      // Object path of wmi namespace
        _bstr_t( username_str),      // User name. NULL = current
        _bstr_t( password_str),      // User password. NULL = current
        0   ,                        // Local.  NULL = current
        NULL,                        // Security flags.
        _bstr_t( upn_str     ),      // Authority (for example, Kerberos) RJM-FixMe?
        0   ,                        // Contect object
        &(self->default_svc)         // pointer to IWbemServices proxy
      );
  if( FAILED( hres))
  {
    PyErr_Format( PyExc_RuntimeError, "Failed to create default_svc IWbemLocator object for namespace: [%s]. Error Code: [0x%x](%ld)", PyString_AsString( self->host), (LONG )hres, (LONG )hres);
    goto PYWMI_CONNECT_ERROR;
  }
  pLoc->Release();


  retval = Py_BuildValue( "i", 1);

PYWMI_CONNECT_ERROR:
  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-  Description
//-  - Creates File at 'remote_file_path' and returns a Windows HANDLE incapsulated
//-     within a PyObject
//-
//-  Arguments
//-  - remote_file_path (required): the fully qualified path of where to create file
//-    - Example: \\MyHost\C$\temp\my_file.txt
//-
//-  Return Value
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value (incapsulated Windows HANDLE)  [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_file_create( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;

  char *kwlist[] = { "remote_file_path", "access", "disposition", NULL };

  DWORD     rval           = 0;
  HANDLE    handle         = 0;
  PyObject *py_file_path   = 0;
  char     *path           = 0;
  DWORD     access         = GENERIC_READ | GENERIC_WRITE;
  DWORD     disposition    = CREATE_ALWAYS;

  rval = PyArg_ParseTupleAndKeywords( args, kwargs, "S|kk", kwlist, &(py_file_path),
                                                                      &(access      ),
                                                                      &(disposition ));
  if( !rval)
  {
    goto PYWMI_FILE_CREATE_ERROR;
  }

  path = PyString_AsString( py_file_path);

  if( (disposition < CREATE_NEW) || (disposition > TRUNCATE_EXISTING))
  {
    PyErr_Format( PyExc_ValueError, "for argument 'disposition' (pos 3)");
    goto PYWMI_FILE_CREATE_ERROR;
  }

  //handle = CreateFile( path, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, CREATE_ALWAYS, 0, NULL);
  handle = CreateFile( path, access, 0, NULL, disposition, 0, NULL);
  if( handle == INVALID_HANDLE_VALUE)
  {
    DWORD le = GetLastError();

    LPVOID lpMsgBuf;
    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM     |
        FORMAT_MESSAGE_IGNORE_INSERTS  ,
        NULL,
        le,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR) &lpMsgBuf,
        0, NULL );

    PyErr_Format( PyExc_RuntimeError, "Error Code: [0x%x](%d) %s", rval, rval, (LPTSTR *)lpMsgBuf);
    LocalFree(lpMsgBuf);
  }
  else
  {
    retval = Py_BuildValue( "l", handle);
  }

PYWMI_FILE_CREATE_ERROR:
  return( retval);

}


//////////////////////////////////////////////////////////////////////////////-
//-  Description
//-  - Writes data to incapulated Windows HANDLE provided by PyWMI_file_create
//-
//-  Arguments
//-  - PyObject (required): incapsulated HANDLE returned by PyWMI_file_create
//-  - PyObject (required): data to be written
//-
//-  Return Value
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value (bytes actually written) [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_file_write ( PyWMIObject *self , PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;
  if( self)
  {
    char *kwlist[] = { "handle", "data", NULL };

    PyObject *py_handle = 0;
    PyObject *py_data   = 0;
    HANDLE    handle    = 0;

    if( !PyArg_ParseTupleAndKeywords( args, kwargs, "OO", kwlist, &(py_handle),
                                                                &(py_data)))
    {
      goto PYWMI_FILE_WRITE_ERROR;
    }

    if( py_handle && py_data)
    {
      handle = (HANDLE)(PyInt_AsLong( py_handle));
      int   rval = 0;
      int   size = PyByteArray_Size( py_data);
      char *data = PyString_AsString( py_data);
      WriteFile( handle, data, size, (LPDWORD)&rval, 0);

      retval = Py_BuildValue( "i", rval);
    }
  }

PYWMI_FILE_WRITE_ERROR:
  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-  Description
//-  - Closes incapulated Windows HANDLE provided by PyWMI_file_create
//-
//-  Arguments
//-  - PyObject (required): incapsulated HANDLE returned by PyWMI_file_create
//-
//-  Return Value
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_file_close ( PyWMIObject *self , PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;

  if( self)
  {
    char *kwlist[] = { "handle", NULL };

    PyObject *py_handle = 0;
    HANDLE    handle    = 0;

    if( !PyArg_ParseTupleAndKeywords( args, kwargs, "O", kwlist, &(py_handle)))
    {
      goto PYWMI_FILE_CLOSE_ERROR;
    }
    if( py_handle)
    {
      handle = (HANDLE)(PyInt_AsLong( py_handle));
      int rval = CloseHandle( handle);
      retval = Py_BuildValue( "i", rval);
    }
  }

PYWMI_FILE_CLOSE_ERROR:
  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-  Description
//-  - Executes (via cmd.exe) 'command' on connected host
//-
//-  Arguments
//-  - PyObject (required): Command to execute on remote host
//-
//-  Return Value
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value (ReturnValue of exec'd process) [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_file_exec( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;

  char *kwlist[] = { "command", "rwd", NULL };

  char     *cmd_prefix = "cmd.exe /Q /c";
  PyObject *py_command = 0;
  PyObject *py_rwd     = 0;
  HANDLE    handle     = 0;
  char     *command    = 0;
  char     *rwd        = 0;

  HRESULT hres = 0;
  //- set up to call the Win32_Process::Create method
  BSTR              MethodName = SysAllocString( L"Create"       );
  BSTR              ClassName  = SysAllocString( L"Win32_Process");
  IWbemClassObject* pClass              = NULL;
  IWbemClassObject* pInParamsDefinition = NULL;
  IWbemClassObject* pClassInstance      = NULL;
  VARIANT           varCommand          = {0};
  VARIANT           varCurrentDirectory = {0};
  IWbemClassObject* pOutParams          = NULL;


  if( !PyArg_ParseTupleAndKeywords( args, kwargs, "S|S", kwlist, &(py_command)))
  {
    goto PYWMI_FILE_EXEC_ERROR;
  }

  hres = self->cimv2_svc->GetObject( ClassName, 0, NULL, &pClass, NULL);
  hres = pClass->GetMethod( MethodName, 0, &pInParamsDefinition, NULL);
  hres = pInParamsDefinition->SpawnInstance( 0, &pClassInstance);

  if( py_command)
  {
    command = PyString_AsString( py_command);
    //- rjm note:  The following two (2) lines may be removed
    //-  need to test rundll32.exe running without cmd_prefix
    py_command = PyString_FromFormat( "%s %s", cmd_prefix, command);
    command = PyString_AsString( py_command);

    //- Create the values for the in parameters
    varCommand.vt = VT_BSTR;
    varCommand.bstrVal = _bstr_t( command );
    //- Store the value for the in parameters
    hres = pClassInstance->Put( L"CommandLine", 0, &varCommand, 0);
  }
  if( py_rwd)
  {
    rwd = PyString_AsString( py_rwd);
    //- Create the values for the in parameters
    varCurrentDirectory.vt = VT_BSTR;
    varCurrentDirectory.bstrVal = _bstr_t( rwd );
    //- Store the value for the in parameters
    hres = pClassInstance->Put( L"CurrentDirectory", 0, &varCurrentDirectory, 0);
  }

  // Execute Method
  hres = self->cimv2_svc->ExecMethod( ClassName, MethodName, 0, NULL, pClassInstance,
                                   &pOutParams,NULL);
  if (FAILED(hres))
  {
    PyErr_Format( PyExc_RuntimeError, "Could not execute method. Error Code: [0x%x](%ld)", hres, hres);
  }
  else
  {
    retval = PyDict_New();

    VARIANT var_retval;
    hres = pOutParams->Get( _bstr_t(L"ReturnValue"), 0, &var_retval, NULL, 0);
    PyDict_SetItem( retval, PyString_FromString( "ReturnValue"),
                    Py_BuildValue( "i", var_retval.lVal));
    if( var_retval.lVal == 0)
    {
      VariantClear(&var_retval);
      hres = pOutParams->Get( _bstr_t(L"ProcessId"), 0, &var_retval, NULL, 0);

      PyDict_SetItem( retval, PyString_FromString( "ProcessId"),
                      Py_BuildValue( "i", var_retval.uintVal));
    }
  }


PYWMI_FILE_EXEC_ERROR:
  SysFreeString(ClassName);
  SysFreeString(MethodName);
  if( pClass)
  {
    pClass->Release();
  }
  if( pInParamsDefinition)
  {
    pInParamsDefinition->Release();
  }
  if( pClassInstance)
  {
    pClassInstance->Release();
  }
  VariantClear(&varCommand);
  if( pOutParams)
  {
    pOutParams->Release();
  }

  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-  Description
//-  - Wrapper for GetExpandedStringValue method of the StdRegProv class
//-
//-  Arguments
//-  - PyObject (required): hKey (uint32)
//-  - PyObject (required): subkey_name
//-  - PyObject (required): value_name
//-
//-  Return Value
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: PyDict object containing the 'ReturnValue' of the ExecMethod
//-    -        :   'sValue' containing the value of hKey/subkey_name/value_name
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_reg_getexpandedstringvalue( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;
  char *kwlist[] = { "hkey", "subkey_name", "value_name", NULL };

  PyObject *py_hkey   = 0;
  PyObject *py_subkey = 0;
  PyObject *py_value  = 0;
  unsigned int  hkey   = 0;
  char         *subkey = 0;
  char         *value  = 0;

  HRESULT hres = 0;
  //- set up to call the Win32_Process::Create method
  BSTR              MethodName = SysAllocString( L"GetExpandedStringValue");
  BSTR              ClassName  = SysAllocString( L"StdRegProv"            );
  IWbemClassObject* pClass              = NULL;
  IWbemClassObject* pInParamsDefinition = NULL;
  IWbemClassObject* pClassInstance      = NULL;
  VARIANT           var_hkey            = {0};
  VARIANT           var_subkey          = {0};
  VARIANT           var_value           = {0};
  IWbemClassObject* pOutParams          = NULL;

  if( !PyArg_ParseTupleAndKeywords( args, kwargs, "ISS", kwlist, &(hkey     ),
                                                                 &(py_subkey),
                                                                 &(py_value )))
  {
    goto PYWMI_REG_GESV_END;
  }

  if( py_subkey )
  {
    subkey = PyString_AsString( py_subkey);
  }
  if( py_value )
  {
    value = PyString_AsString( py_value);
  }

  hres = self->default_svc->GetObject( ClassName, 0, NULL, &pClass, NULL);
  hres = pClass->GetMethod( MethodName, 0, &pInParamsDefinition, NULL);
  hres = pInParamsDefinition->SpawnInstance( 0, &pClassInstance);

  //- Build inparams
  //- VARIANT hkey
  var_hkey.vt        = VT_UINT;
  var_hkey.uintVal   = hkey   ;
  //- VARIANT subkey
  var_subkey.vt      = VT_BSTR;
  var_subkey.bstrVal = _bstr_t( subkey);
  //- VARIANT value
  var_value.vt       = VT_BSTR;
  var_value.bstrVal  = _bstr_t( value);

  //- put the variant values in the class instance
  hres = pClassInstance->Put( L"hDefKey"    , 0, &var_hkey  , 0);
  hres = pClassInstance->Put( L"sSubKeyName", 0, &var_subkey, 0);
  hres = pClassInstance->Put( L"sValueName" , 0, &var_value , 0);

  //- Execute Method
  hres = self->default_svc->ExecMethod( ClassName, MethodName, 0, NULL, pClassInstance,
                                       &pOutParams,NULL);
  if (FAILED(hres))
  {
    PyErr_Format( PyExc_RuntimeError,
           "Could not execute method: [%s]. Error Code: [0x%x](%ld)", MethodName, hres, hres);
  }
  else
  {
    retval = PyDict_New();

    VARIANT var_retval;
    hres = pOutParams->Get( _bstr_t(L"ReturnValue"), 0, &var_retval, NULL, 0);
    PyDict_SetItem( retval, PyString_FromString( "ReturnValue"), Py_BuildValue( "i", var_retval.lVal));
    if( var_retval.lVal == 0)
    {
      VariantClear(&var_retval);
      hres = pOutParams->Get( _bstr_t(L"sValue"), 0, &var_retval, NULL, 0);

      PyDict_SetItem( retval, PyString_FromString( "sValue"), Py_BuildValue( "u", var_retval.bstrVal));
    }
    else
    {
      PyDict_SetItem( retval, PyString_FromString( "sValue"), Py_BuildValue( "u", L""));
    }
  }

PYWMI_REG_GESV_END:
  SysFreeString(ClassName);
  SysFreeString(MethodName);
  if( pClass)
  {
    pClass->Release();
  }
  if( pInParamsDefinition)
  {
    pInParamsDefinition->Release();
  }
  if( pClassInstance)
  {
    pClassInstance->Release();
  }
  VariantClear(&var_hkey  );
  VariantClear(&var_subkey);
  VariantClear(&var_value );
  if( pOutParams)
  {
    pOutParams->Release();
  }

  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   'disconnects' from remote wmi services, releases IWbemServices objects
//-
//- Arguments:
//-  void
//-
//- Return Value:
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_disconnect( PyWMIObject *self)
{
  PyObject *retval = 0;

  if( self->cimv2_svc)
  {
    self->cimv2_svc->Release();
  }
  if( self->default_svc)
  {
    self->default_svc->Release();
  }
  retval = Py_BuildValue( "i", 1);

  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Wrapper for Windows 'DeleteFile'
//-
//- Arguments:
//-  - PyObject (required): remote_file_path
//-
//- Return Value:
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_file_delete( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;

  char *kwlist[] = { "remote_file_path", NULL };

  PyObject *py_remote_file_path = 0;
  char     *remote_file_path    = 0;

  if( !PyArg_ParseTupleAndKeywords( args, kwargs, "S", kwlist, &(py_remote_file_path)))
  {
    goto PYWMI_FILE_DELETE_ERROR;
  }
  remote_file_path = PyString_AsString( py_remote_file_path);

  DWORD rval = DeleteFile( remote_file_path );
  if( rval == 0)
  {
    DWORD le = GetLastError();

    LPVOID lpMsgBuf;
    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM     |
        FORMAT_MESSAGE_IGNORE_INSERTS  ,
        NULL,
        le,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR) &lpMsgBuf,
        0, NULL );

    PyErr_Format( PyExc_RuntimeError, "Error Code: [0x%x](%d) %s", rval, rval, lpMsgBuf);

    LocalFree(lpMsgBuf);
  }
  else
  {
    retval = Py_BuildValue( "k", rval);
  }

PYWMI_FILE_DELETE_ERROR:

  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//-
//- Description:
//-   Wrapper for Windows 'CopyFile'
//-
//- Arguments:
//-  - PyObject (required): local_file_path
//-  - PyObject (required): remote_file_path
//-
//- Return Value:
//-  - PyObject *
//-    - Failure: 0 (NULL) [will result in thrown python exception]
//-    - Success: Non-Zero value [via Py_BuildValue]
//////////////////////////////////////////////////////////////////////////////-
static PyObject *
PyWMI_file_get( PyWMIObject *self, PyObject *args, PyObject *kwargs )
{
  PyObject *retval = 0;

  char *kwlist[] = { "local_file_path", "remote_file_path", NULL };

  PyObject *py_local_file_path  = 0;
  PyObject *py_remote_file_path = 0;

  char     *local_file_path     = 0;
  char     *remote_file_path    = 0;

  if( !PyArg_ParseTupleAndKeywords( args, kwargs, "SS", kwlist, &(py_local_file_path),
                                                              &(py_remote_file_path)))
  {
    goto PYWMI_FILE_GET_ERROR;
  }
  local_file_path  = PyString_AsString( py_local_file_path);
  remote_file_path = PyString_AsString( py_remote_file_path);

  DWORD rval = CopyFile( remote_file_path, local_file_path, false);
  if( rval == 0)
  {
    rval = GetLastError();

    LPVOID lpMsgBuf;
    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM     |
        FORMAT_MESSAGE_IGNORE_INSERTS  ,
        NULL,
        rval,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR) &lpMsgBuf,
        0, NULL );

    PyErr_Format( PyExc_RuntimeError, "Error Code: [0x%x](%d) %s", rval, rval, lpMsgBuf);
    LocalFree(lpMsgBuf);

  }
  else
  {
    retval = Py_BuildValue( "k", rval);
  }

PYWMI_FILE_GET_ERROR:
  return( retval);
}


//////////////////////////////////////////////////////////////////////////////-
//- EOF
//////////////////////////////////////////////////////////////////////////////-
