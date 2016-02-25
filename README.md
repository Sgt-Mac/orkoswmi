## Synopsis

OrkosWMI is a thin WMI shim intended for use with root9B's [Orkos - Credential Assessment Capability](https://www.root9b.com/products).

## Requirements

`Orkos`

## Build Requirements

* Modern version of Windows
* Python 2.7 (32/64 bit)
* pip (Python package management system)
* [Microsoft Visual C++ Compiler for Python 2.7](https://download.microsoft.com/download/7/9/6/796EF2E4-801B-4FC4-AB28-B59FBF6D907B/VCForPython27.msi)

## Install steps

* pip install git+https://github.com/root9b/orkoswmi.git


## Usage

```
import orkoswmi

owmi = orkoswmi.PyWMI() # optional orkoswmi.PyWMI( 'hostname')

# do_krb: user kerberos authentication vs NTLM
# owmi.connect( username='username', password='password', do_krb=False)):
if 1 == owmi.connect( 'hostname-or-addr', 'username', 'password', do_krb=True)):
  owmi.mount_share( 'share_full_path') # optional
  # ... do some stuff ...
  owmi.unmount_share()
  owmi.disconnect()

```
