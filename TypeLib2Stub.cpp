// TypeLib2Stub.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

void help() {
	std::cout <<
		"Usage: " << std::endl <<
		"\tTypeLib2Stub <typelib_file>" << std::endl;
}

std::map<VARTYPE, std::wstring> _var_type_name;

void describe_type(std::set<std::wstring> &fwdRefs, std::set<std::wstring> &knownTypes, TYPEDESC *ptdesc, std::wostream &output, ITypeInfo *tinfo) {
	auto indirectionLevel = 0;
	while (ptdesc->vt == VT_PTR) {
		indirectionLevel++;
		ptdesc = ptdesc->lptdesc;
	}
	if (ptdesc->vt == VT_USERDEFINED) {
		ITypeInfo *refTinfo;
		tinfo->GetRefTypeInfo(ptdesc->hreftype, &refTinfo);
		BSTR name;
		refTinfo->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL);
		TYPEATTR *pRefAttr;
		refTinfo->GetTypeAttr(&pRefAttr);

		if (pRefAttr->typekind == TKIND_ALIAS) {
			// remove the alias, use the pointed-to type
			describe_type(fwdRefs, knownTypes, &pRefAttr->tdescAlias, output, tinfo);
		}
		else {
			output << name;
			if (knownTypes.find(name) == knownTypes.end()) {
				if (pRefAttr->typekind == TKIND_ENUM) {
					fwdRefs.emplace(std::wstring(L"enum ") + name);
				}
				else {
					fwdRefs.emplace(std::wstring(L"struct ") + name);
				}
			}
		}
		refTinfo->ReleaseTypeAttr(pRefAttr);
		::SysFreeString(name);
		refTinfo->Release();
	}
	else {
		output << _var_type_name[ptdesc->vt].c_str();
	}
	output << std::wstring(indirectionLevel, L'*').c_str();
}

std::wstring guid_to_tring(GUID &uuid) {
	LPOLESTR olestr;
	::StringFromIID(uuid, &olestr);
	std::wstring ret(olestr);
	::CoTaskMemFree(olestr);
	return ret;
}

bool describe_function(int fn_idx, std::set<ULONG>& knownDispIDs, std::set<std::wstring> &fwdRefs, std::set<std::wstring> &knownTypes, ITypeInfo *tinfo, std::wostream &buff) {
	FUNCDESC *pFuncDesc;
	tinfo->GetFuncDesc(fn_idx, &pFuncDesc);
	BSTR names[20];
	UINT cNames = 20;
	tinfo->GetNames(pFuncDesc->memid, names, cNames, &cNames);
	if (::lstrcmpW(L"url", names[0]) == 0) {
		::lstrcpyW(names[0], L"URL");
	}
	auto ret = false;

	if (knownDispIDs.find(pFuncDesc->memid) == knownDispIDs.end())
	{
		ret = true;
		buff << "\tvirtual HRESULT ";

		if (pFuncDesc->callconv == CC_STDCALL) {
			buff << "__stdcall ";
		}
		auto is_property = false;
		if (pFuncDesc->invkind == INVOKE_PROPERTYGET) {
			buff << "get_";
			is_property = true;
		}
		else if (pFuncDesc->invkind == INVOKE_PROPERTYPUT) {
			buff << "put_";
			is_property = true;
		}
		else if (pFuncDesc->invkind == INVOKE_PROPERTYPUTREF) {
			buff << "put_";
			is_property = true;
		}
		buff << names[0] << L" (";
		for (auto paramidx = 0; paramidx < pFuncDesc->cParams; paramidx++) {
			if (paramidx > 0) {
				buff << ", ";
			}
			auto ptdesc = &pFuncDesc->lprgelemdescParam[paramidx].tdesc;
			describe_type(fwdRefs, knownTypes, ptdesc, buff, tinfo);
			auto name_idx = (paramidx + (is_property ? 0 : 1)) % cNames;
			buff << L" " << names[name_idx];
		}
		if (pFuncDesc->elemdescFunc.tdesc.vt != VT_VOID &&
			pFuncDesc->elemdescFunc.tdesc.vt != VT_HRESULT) {
			if (pFuncDesc->cParams)
				buff << ", ";
			buff << "/*[retval]*/ ";
			describe_type(fwdRefs, knownTypes, &pFuncDesc->elemdescFunc.tdesc, buff, tinfo);
			buff << L" *retval";
		}

		buff << ") " ;
		//pFuncDesc->lprgelemdescParam[]
		for (auto stridx = 0; stridx < cNames; stridx++) {
			::SysFreeString(names[stridx]);
		}
	}
	tinfo->ReleaseFuncDesc(pFuncDesc);
	return ret;
}

std::wstring write_implementation(HREFTYPE reftype, std::set<std::wstring> &implementedInterfaces, ITypeInfo *tinfo, std::set<ULONG> knownDispIDs, std::wostream &output) 
{
	ITypeInfo *reftinfo;
	tinfo->GetRefTypeInfo(reftype, &reftinfo);
	BSTR name;
	reftinfo->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL);
	std::wstring ret(name);
	if (implementedInterfaces.find(name) == implementedInterfaces.end()) {
		// stock implementations
		if (0 == ::lstrcmpW(name, L"IUnknown")) {
			output << L"template <typename T>" << std::endl
				<< L"class IUnknownImpl: public T {" << std::endl
				<< "\tULONG _refcount;" << std::endl
				<< "protected:" << std::endl
				<< "\tvirtual bool IsSupported(const IID& riid) { return riid == __uuidof(IUnknown); }" << std::endl
				<< "public:" << std::endl
				<< "\tIUnknownImpl() : _refcount(1L) {}" << std::endl
				<< "\tvirtual HRESULT __stdcall QueryInterface(const IID &riid, void **ppv) override {" << std::endl
				<< "\t\tif (IsSupported(riid)) { *ppv = this; AddRef(); return S_OK; }" << std::endl
				<< "\t\treturn E_NOINTERFACE;" << std::endl
				<< "\t}" << std::endl
				<< "\tvirtual ULONG __stdcall AddRef() override { return ++_refcount; }" << std::endl
				<< "\tvirtual ULONG __stdcall Release() override { if (0 >= --_refcount) { delete this; return 0; } return _refcount; }" << std::endl
				<< "};" << std::endl;
			implementedInterfaces.emplace(L"IUnknown");
		}
		else if (0 == ::lstrcmpW(name, L"IDispatch")) {
			// first make sure to have the base type implementations
			TYPEATTR *prefattr;
			reftinfo->GetTypeAttr(&prefattr);
			std::set<std::wstring> baseTypes;
			for (auto impl_idx = 0; impl_idx < prefattr->cImplTypes; impl_idx++) {
				HREFTYPE baseref;
				reftinfo->GetRefTypeOfImplType(impl_idx, &baseref);
				baseTypes.emplace(write_implementation(baseref, implementedInterfaces, reftinfo, knownDispIDs, output));
			}
			output << L"template <typename T>" << std::endl
				<< L"class IDispatchImpl: public IUnknownImpl<T> {" << std::endl
				<< "protected:" << std::endl
				<< "\tvirtual bool IsSupported(const IID& riid) override { return riid == __uuidof(IDispatch) || IUnknownImpl<T>::IsSupported(riid); }" << std::endl
				<< "public:" << std::endl
				<< "virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(__RPC__out UINT *pctinfo)"
				<< " { return E_NOTIMPL; }" << std::endl
				<< "virtual HRESULT STDMETHODCALLTYPE GetTypeInfo( UINT iTInfo,LCID lcid, __RPC__deref_out_opt ITypeInfo **ppTInfo)"
				<< " { return E_NOTIMPL; }" << std::endl
				<< "virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames( __RPC__in REFIID riid,__RPC__in_ecount_full(cNames) LPOLESTR *rgszNames, __RPC__in_range(0, 16384) UINT cNames,LCID lcid, __RPC__out_ecount_full(cNames) DISPID *rgDispId)"
				<< " { return E_NOTIMPL; }" << std::endl
				<< "virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke(_In_  DISPID dispIdMember,_In_  REFIID riid,_In_  LCID lcid,_In_  WORD wFlags,_In_  DISPPARAMS *pDispParams,_Out_opt_  VARIANT *pVarResult,_Out_opt_  EXCEPINFO *pExcepInfo,_Out_opt_  UINT *puArgErr)"
				<< " { return E_NOTIMPL; }" << std::endl
				<< "};" << std::endl;
			implementedInterfaces.emplace(L"IDispatch");
		}
		else {
			// first make sure to have the base type implementations
			TYPEATTR *prefattr;
			reftinfo->GetTypeAttr(&prefattr);
			std::set<std::wstring> baseTypes;
			for (auto impl_idx = 0; impl_idx < prefattr->cImplTypes; impl_idx++) {
				HREFTYPE baseref;
				reftinfo->GetRefTypeOfImplType(impl_idx, &baseref);
				baseTypes.emplace(write_implementation(baseref, implementedInterfaces, reftinfo, knownDispIDs, output));
			}

			output
				<< L"template <typename T>" << std::endl
				<< L"class " << name << L"Impl";
			if (baseTypes.size()) {
				output << L":";
				auto writtensome = false;
				for (auto btype : baseTypes) {
					if (writtensome)
						output << L", ";
					else
						writtensome = true;
					output << L" public " << btype << L"Impl<T>";
				}
			}
			output << L" {" << std::endl;
			output << "protected:" << std::endl
				<< "\tvirtual bool IsSupported(const IID& riid) override { return riid == __uuidof("
				<< name
				<< ") ";
			for (auto btype : baseTypes) {
				output << "|| "<< btype << "Impl<T>::IsSupported(riid)";
			}
			output << "; }" << std::endl
				<< "public:" << std::endl;
			std::set<std::wstring> fwdRefs, knownTypes;
			for (auto idx = 0; idx < prefattr->cFuncs; idx++) {
				if (describe_function(idx, knownDispIDs, fwdRefs, knownTypes, reftinfo, output))
					output
					<< (prefattr->typekind == TKIND_DISPATCH ? "" : " override")
					<< L" {\n\t\treturn E_NOTIMPL;\n\t}" << std::endl;
			}
			output << L"};" << std::endl;

			implementedInterfaces.emplace(name);
			reftinfo->ReleaseTypeAttr(prefattr);
		}
	}
	::SysFreeString(name);
	reftinfo->Release();
	return ret;
}


int main(int argc, const char *argv[])
{
	if (argc < 2) {
		help();
		return 1;
	}

	_var_type_name[VT_EMPTY] = L"?";
	_var_type_name[VT_NULL] = L"NULL";
	_var_type_name[VT_I2] = L"short";
	_var_type_name[VT_I4] = L"long";
	_var_type_name[VT_R4] = L"float";
	_var_type_name[VT_R8] = L"double";
	_var_type_name[VT_CY] = L"currency?";
	_var_type_name[VT_DATE] = L"date?";
	_var_type_name[VT_BSTR] = L"BSTR";
	_var_type_name[VT_DISPATCH] = L"IDispatch*";
	_var_type_name[VT_ERROR] = L"HRESULT";
	_var_type_name[VT_BOOL] = L"VARIANT_BOOL";
	_var_type_name[VT_VARIANT] = L"VARIANT";
	_var_type_name[VT_UNKNOWN] = L"IUnknown*";
	_var_type_name[VT_DECIMAL] = L"DECIMAL???";
	_var_type_name[VT_I1] = L"char";
	_var_type_name[VT_UI1] = L"unsigned char";
	_var_type_name[VT_UI2] = L"unsigned short";
	_var_type_name[VT_UI4] = L"unsigned long";
	_var_type_name[VT_I8] = L"long long";
	_var_type_name[VT_UI8] = L"unsigned long long";
	_var_type_name[VT_INT] = L"int";
	_var_type_name[VT_UINT] = L"unsigned int";
	_var_type_name[VT_VOID] = L"void";
	_var_type_name[VT_HRESULT] = L"HRESULT";
	_var_type_name[VT_PTR] = L"void *";

	std::set<std::wstring> knownTypes;
	knownTypes.emplace(L"IXMLDOMImplementation");
	knownTypes.emplace(L"IXMLDOMNode");
	knownTypes.emplace(L"tagDOMNodeType");
	knownTypes.emplace(L"IXMLDOMNodeList");
	knownTypes.emplace(L"IXMLDOMNamedNodeMap");
	knownTypes.emplace(L"IXMLDOMDocument");
	knownTypes.emplace(L"IXMLDOMDocumentType");
	knownTypes.emplace(L"IXMLDOMElement");
	knownTypes.emplace(L"IXMLDOMAttribute");
	knownTypes.emplace(L"IXMLDOMDocumentFragment");
	knownTypes.emplace(L"IXMLDOMText");
	knownTypes.emplace(L"IXMLDOMCharacterData");
	knownTypes.emplace(L"IXMLDOMComment");
	knownTypes.emplace(L"IXMLDOMCDATASection");
	knownTypes.emplace(L"IXMLDOMProcessingInstruction");
	knownTypes.emplace(L"IXMLDOMEntityReference");
	knownTypes.emplace(L"IXMLDOMParseError");
	knownTypes.emplace(L"DISPPARAMS");
	knownTypes.emplace(L"EXCEPINFO");
	knownTypes.emplace(L"GUID");
	knownTypes.emplace(L"IXMLDOMNotation");
	knownTypes.emplace(L"IXMLDOMEntity");
	knownTypes.emplace(L"IXTLRuntime");
	knownTypes.emplace(L"IXMLElementCollection");
	knownTypes.emplace(L"IXMLDocument");
	knownTypes.emplace(L"IXMLElement");
	knownTypes.emplace(L"IXMLDocument2");
	knownTypes.emplace(L"IXMLElement2");
	knownTypes.emplace(L"IXMLAttribute");
	knownTypes.emplace(L"IXMLError");
	knownTypes.emplace(L"_xml_error");
	knownTypes.emplace(L"tagXMLEMEM_TYPE");
	knownTypes.emplace(L"XMLDOMDocumentEvents");

	std::set<ULONG> knownDispIDs;
	knownDispIDs.emplace(0x60000000); // IUnknown
	knownDispIDs.emplace(0x60000001);
	knownDispIDs.emplace(0x60000002);
	knownDispIDs.emplace(0x60010000); // IDispatch
	knownDispIDs.emplace(0x60010001);
	knownDispIDs.emplace(0x60010002);
	knownDispIDs.emplace(0x60010003);
	knownDispIDs.emplace(0x60010004);

	std::set<std::wstring> implementedInterfaces;
	std::set<std::wstring> coclasses;
	std::map<std::wstring, std::set<std::wstring>> inheritance;

	std::wstring wname(argv[1], argv[1] + ::lstrlenA(argv[1]));
	ITypeLib *tlib;
	auto hr = ::LoadTypeLib(wname.c_str(), &tlib);
	if (SUCCEEDED(hr)) {
		std::wcout << "// generated for " << wname.c_str() << std::endl <<
			"#include <windows.h>" << std::endl;

		std::wcout << std::endl;

		auto count = tlib->GetTypeInfoCount();

		for (auto idx = 0; idx < count; idx++) {
			ITypeInfo *tinfo;
			hr = tlib->GetTypeInfo(idx, &tinfo);
			if (SUCCEEDED(hr)) {
				BSTR name;
				BSTR doc;
				tinfo->GetDocumentation(MEMBERID_NIL, &name, &doc, NULL, NULL);

				TYPEATTR *attr;
				tinfo->GetTypeAttr(&attr);
				if (knownTypes.find(name) == knownTypes.end()) {
					knownTypes.emplace(name);
					std::wstringstream buff;

					std::set<std::wstring> fwdRefs;
					std::set<HREFTYPE> requiredInterfaces;

					auto derived_start_offset = 0;
					auto close_braces = true;

					if (attr->typekind == TKIND_DISPATCH || attr->typekind == TKIND_INTERFACE || attr->typekind == TKIND_RECORD) {
						buff << L" /* " << (doc ? doc : L"(no documentation)") << L" */" << std::endl;
						buff << L"struct ";
						if (attr->guid != GUID_NULL) {
							buff << "__declspec(uuid(\"" <<
								guid_to_tring(attr->guid).c_str()
								<< "\"))" << std::endl;
						}
						buff << name;
						if (attr->cImplTypes) {
							buff << ":";
							for (auto idximpl = 0; idximpl < attr->cImplTypes; idximpl++) {
								if (idximpl > 0) {
									buff << ",";
								}
								HREFTYPE reftype;
								tinfo->GetRefTypeOfImplType(idximpl, &reftype);
								ITypeInfo *reftinfo;
								tinfo->GetRefTypeInfo(reftype, &reftinfo);
								BSTR basename;
								reftinfo->GetDocumentation(MEMBERID_NIL, &basename, NULL, NULL, NULL);
								inheritance[name].emplace(basename);
								buff << L" public " << basename;
								TYPEATTR *pBaseAttr;
								reftinfo->GetTypeAttr(&pBaseAttr);
								// derived_start_offset = pBaseAttr->cFuncs;
								reftinfo->ReleaseTypeAttr(pBaseAttr);
								reftinfo->Release();
							}
						}
						buff << std::endl;
						buff << "{" << std::endl;
					}
					else if (attr->typekind == TKIND_COCLASS) {
						coclasses.emplace(name);
						buff << L" /* " << (doc ? doc : L"(no documentation)") << L" */" << std::endl;
						buff << L"class ";
						if (attr->guid != GUID_NULL) {
							buff << "__declspec(uuid(\"" <<
								guid_to_tring(attr->guid).c_str()
								<< "\"))" << std::endl;
						}
						buff << L"Stub" << name;
						if (attr->cImplTypes) {
							buff << ":";
							for (auto idximpl = 0; idximpl < attr->cImplTypes; idximpl++) {
								if (idximpl > 0) {
									buff << ",";
								}
								HREFTYPE reftype;
								tinfo->GetRefTypeOfImplType(idximpl, &reftype);
								ITypeInfo *reftinfo;
								tinfo->GetRefTypeInfo(reftype, &reftinfo);
								BSTR basename;
								reftinfo->GetDocumentation(MEMBERID_NIL, &basename, NULL, NULL, NULL);
								inheritance[name].emplace(basename);
								buff << L" public " << basename << L"Impl<" << basename << L">";
								requiredInterfaces.emplace(reftype);
								reftinfo->Release();
							}
						}
						buff << std::endl;
						buff << "{" << std::endl;
						buff << "public:" << std::endl;
					}
					else if (attr->typekind == TKIND_ENUM) {
						buff << L" /* " << (doc ? doc : L"(no documentation)") << L" */" << std::endl;
						buff << L"enum ";
						if (attr->guid != GUID_NULL) {
							buff << "__declspec(uuid(\"" <<
								guid_to_tring(attr->guid).c_str()
								<< "\"))" << std::endl;
						}
						buff << name;
						if (attr->cImplTypes) {
							buff << ":";
							for (auto idximpl = 0; idximpl < attr->cImplTypes; idximpl++) {
								if (idximpl > 0) {
									buff << ",";
								}
								HREFTYPE reftype;
								tinfo->GetRefTypeOfImplType(idximpl, &reftype);
								ITypeInfo *reftinfo;
								tinfo->GetRefTypeInfo(reftype, &reftinfo);
								BSTR basename;
								reftinfo->GetDocumentation(MEMBERID_NIL, &basename, NULL, NULL, NULL);
								inheritance[name].emplace(basename);
								buff << L" public " << basename;
								reftinfo->Release();
							}
						}
						buff << std::endl;
						buff << "{" << std::endl;
					}
					else if (attr->typekind == TKIND_ALIAS) {
						buff << "typedef ";
						describe_type(fwdRefs, knownTypes, &attr->tdescAlias, buff, tinfo);
						buff << L" " << name << ";" << std::endl;
						close_braces = false;
					}
					else {
						buff << L" /* " << (doc ? doc : L"(no documentation)") << L" */" << std::endl;
					}

					for (auto fn_idx = derived_start_offset; fn_idx < attr->cFuncs; fn_idx++) {
						if (describe_function(fn_idx, knownDispIDs, fwdRefs, knownTypes, tinfo, buff))
							buff << L" = 0;" << std::endl;
					}

					for (auto v_idx = 0; v_idx < attr->cVars; v_idx++) {
						VARDESC *pVarDesc;
						tinfo->GetVarDesc(v_idx, &pVarDesc);
						LPOLESTR var_name;
						UINT dummy;
						if (SUCCEEDED(tinfo->GetNames(pVarDesc->memid, &var_name, 1, &dummy))) {
							buff << var_name;
							if (pVarDesc->varkind == VAR_CONST) {
								buff << L" = " << pVarDesc->lpvarValue->intVal;
							}
							buff << L", " << std::endl;
							::SysFreeString(var_name);
						}
						tinfo->ReleaseVarDesc(pVarDesc);
					}

					if (close_braces)
						buff << "};" << std::endl << std::endl;

					for (auto fwdr : fwdRefs) {
						std::wcout << fwdr.c_str() << L";" << std::endl;
					}

					for (auto reqIface : requiredInterfaces) {
						write_implementation(reqIface, implementedInterfaces, tinfo, knownDispIDs, std::wcout);
					}

					std::wcout << buff.str().c_str();
				}
				if (doc) ::SysFreeString(doc);
				if (name) ::SysFreeString(name);
				tinfo->ReleaseTypeAttr(attr);
				tinfo->Release();
			}
		}
		tlib->Release();
	}
	else {
		std::wcout << L"Unable to open " <<
			wname.c_str() <<
			L" error " << hr;
	}
    return 0;
}

