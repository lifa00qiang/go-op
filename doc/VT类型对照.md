| VT 常量              | Go 类型            | 描述              |    |
|--------------------|------------------|-----------------|----|
| VT_EMPTY           | 	nil             | 	空值             |    |
| VT_NULL            | 	nil	            | NULL 值          |    |
| VT_I2              | 	int16	          | 16 位有符号整数       |
| VT_I4              | 	int32	          | 32 位有符号整数       |
| VT_R4              | 	float32         | 32 位浮点数         |
| VT_R8              | 	float64         | 	64 位浮点数        |
| VT_CY              | 	int64	          | 货币类型（64 位整数）    |
| VT_DATE            | 	time.Time       | 	日期时间类型         |
| VT_BSTR            | 	string	         | 宽字符字符串          |
| VT_DISPATCH        | 	*IDispatch      | 	COM 对象指针       |
| VT_ERROR           | 	uint32          | 	错误码            |
| VT_BOOL            | 	bool	           | 布尔值             |
| VT_VARIANT         | 	VARIANT         | 	VARIANT 类型     |
| VT_UNKNOWN         | 	*IUnknown	      | COM 对象指针        |
| VT_DECIMAL         | 	decimal.Decimal | 	十进制数           |
| VT_I1              | 	int8	           | 8 位有符号整数        |
| VT_UI1             | 	uint8	          | 8 位无符号整数        |
| VT_UI2             | 	uint16	         | 16 位无符号整数       |
| VT_UI4             | 	uint32	         | 32 位无符号整数       |
| VT_I8              | 	int64	          | 64 位有符号整数       |
| VT_UI8             | 	uint64	         | 64 位无符号整数       |
| VT_INT             | 	int             | 	平台相关的有符号整数     |
| VT_UINT            | 	uint	           | 平台相关的无符号整数      |
| VT_VOID            | 	nil             | 	无类型            |
| VT_HRESULT         | 	uint32          | 	HRESULT 类型     |
| VT_PTR             | 	unsafe.Pointer  | 	指针类型           |
| VT_SAFEARRAY       | 	*SafeArray	     | SAFEARRAY 类型    |
| VT_CARRAY          | 	unsafe.Pointer  | 	C 风格数组         |
| VT_USERDEFINED     | 	interface{}     | 	用户定义类型         |
| VT_LPSTR           | 	string	         | ANSI 字符串        |
| VT_LPWSTR          | 	string	         | 宽字符字符串          |
| VT_RECORD          | 	struct{}	       | 记录类型            |
| VT_INT_PTR         | 	uintptr         | 	平台相关的指针大小整数    |
| VT_UINT_PTR        | 	uintptr         | 	平台相关的指针大小无符号整数 |
| VT_FILETIME        | 	time.Time	      | 文件时间            |
| VT_BLOB            | 	[]byte          | 	二进制数据          |
| VT_STREAM          | 	*IStream	       | 流对象             |
| VT_STORAGE         | 	*IStorage       | 	存储对象           |
| VT_STREAMED_OBJECT | 	*IStream	       | 流式对象            |
| VT_STORED_OBJECT   | 	*IStorage       | 	存储对象           |
| VT_BLOB_OBJECT     | 	[]byte          | 	二进制对象          |
| VT_CF              | 	*CLIPDATA       | 	剪贴板格式          |
| VT_CLSID           | 	*GUID	          | GUID 类型         |
| VT_BSTR_BLOB       | 	string	         | BSTR 二进制数据      |
| VT_VECTOR          | 	[]interface{}   | 	向量类型           |
| VT_ARRAY           | 	*SafeArray      | 	数组类型           |
| VT_BYREF           | 	unsafe.Pointer  | 	引用类型           |
| VT_RESERVED        | 	uint16	         | 保留类型            |
| VT_ILLEGAL         | 	uint16	         | 非法类型            |
| VT_ILLEGALMASKED   | 	uint16	         | 非法类型掩码          |
| VT_TYPEMASK        | 	uint16	         | 类型掩码            |