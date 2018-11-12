// Helper for converting SafeArray to array of objects.

package ole

import (
	"log"
	"unsafe"
)

type SafeArrayConversion struct {
	Array *SafeArray
}

func (sac *SafeArrayConversion) ToStringArray() (strings []string) {
	totalElements, _ := sac.TotalElements(0)
	strings = make([]string, totalElements)

	for i := int32(0); i < totalElements; i++ {
		strings[int32(i)], _ = safeArrayGetElementString(sac.Array, i)
	}

	return
}

func (sac *SafeArrayConversion) ToByteArray() (bytes []byte) {
	totalElements, _ := sac.TotalElements(0)
	bytes = make([]byte, totalElements)

	for i := int32(0); i < totalElements; i++ {
		safeArrayGetElement(sac.Array, i, unsafe.Pointer(&bytes[int32(i)]))
	}

	return
}

func (sac *SafeArrayConversion) ToValueArray2() (values [][]interface{}) {
	totalElements1, _ := sac.TotalElements(2)
	totalElements2, _ := sac.TotalElements(1)
	te1, te2 := int(totalElements1), int(totalElements2)
	vt, _ := safeArrayGetVartype(sac.Array)

	log.Println(totalElements1, totalElements2)
	values = make([][]interface{}, te1)
	for j := 0; j < te1; j++ {
		row := make([]interface{}, te2)
		for i := 0; i < te2; i++ {
			idx := [2]int64{int64(i), int64(j)}
			idxPt := unsafe.Pointer(&idx)
			switch VT(vt) {
			case VT_BOOL:
				var v bool
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_I1:
				var v int8
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_I2:
				var v int16
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_I4:
				var v int32
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_I8:
				var v int64
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_UI1:
				var v uint8
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_UI2:
				var v uint16
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_UI4:
				var v uint32
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_UI8:
				var v uint64
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_R4:
				var v float32
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_R8:
				var v float64
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_BSTR:
				var v string
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v
			case VT_VARIANT:
				var v VARIANT
				sac.Array.GetElement(idxPt, unsafe.Pointer(&v))
				row[i] = v.Value()
			default:
				// TODO
			}
		}
		values[j] = row
	}
	return
}

func (sac *SafeArrayConversion) ToValueArray() (values []interface{}) {
	totalElements, _ := sac.TotalElements(0)
	values = make([]interface{}, totalElements)
	vt, _ := safeArrayGetVartype(sac.Array)

	for i := int32(0); i < totalElements; i++ {
		switch VT(vt) {
		case VT_BOOL:
			var v bool
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I1:
			var v int8
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I2:
			var v int16
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I4:
			var v int32
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_I8:
			var v int64
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI1:
			var v uint8
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI2:
			var v uint16
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI4:
			var v uint32
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_UI8:
			var v uint64
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_R4:
			var v float32
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_R8:
			var v float64
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_BSTR:
			var v string
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v
		case VT_VARIANT:
			var v VARIANT
			safeArrayGetElement(sac.Array, i, unsafe.Pointer(&v))
			values[i] = v.Value()
		default:
			// TODO
		}
	}

	return
}

func (sac *SafeArrayConversion) GetType() (varType uint16, err error) {
	return safeArrayGetVartype(sac.Array)
}

func (sac *SafeArrayConversion) GetDimensions() (dimensions *uint32, err error) {
	return safeArrayGetDim(sac.Array)
}

func (sac *SafeArrayConversion) GetSize() (length *uint32, err error) {
	return safeArrayGetElementSize(sac.Array)
}

func (sac *SafeArrayConversion) TotalElements(index uint32) (totalElements int32, err error) {
	if index < 1 {
		index = 1
	}

	// Get array bounds
	var LowerBounds int32
	var UpperBounds int32

	LowerBounds, err = safeArrayGetLBound(sac.Array, index)
	if err != nil {
		return
	}

	UpperBounds, err = safeArrayGetUBound(sac.Array, index)
	if err != nil {
		return
	}

	totalElements = UpperBounds - LowerBounds + 1
	return
}

// Release Safe Array memory
func (sac *SafeArrayConversion) Release() {
	safeArrayDestroy(sac.Array)
}
