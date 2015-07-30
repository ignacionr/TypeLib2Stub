# TypeLib2Stub
Creates a single-file C++ stub to use in-lieu of COM classes, in order to support unit testing.

For the time being, you would create a stub implementation like this:
  TypeLib2Stub.exe <path-to-the-typelib-or-dll> > stubwhatever.cpp

Then you can compile like:
  cl -c stubwhatever.cpp /Zi

Then link:
  link stubwhatever.obj /DLL /DEBUG /EXPORT:DllGetClassObject

And you will have a stub implementation compiled!
