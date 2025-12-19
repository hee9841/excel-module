## 스타일 관련 메모 (review2)

- `CustomExcelCellStyle`는 생성자에서 `configure()`를 호출하면 하위 클래스 필드가 아직 준비되지 않은 시점에 추상 메서드가 실행될 수 있습니다. `apply()` 호출 시점에 `ExcelCellStyleConfigurer`를 최초 1회만 준비하도록 지연 초기화하면 필드 초기화 순서를 깨뜨리지 않고, 디버깅도 `apply()` 콜 스택을 통해 더 명확해집니다.

- `ExcelCellStyleConfigurer` 세터들이 `void`라서 연쇄 호출이 불가능합니다. 반환 타입을 `ExcelCellStyleConfigurer`로 바꾸면 `excelColor(...).excelAlign(...).excelBorder(...)`처럼 한눈에 읽히는 플루언트 DSL이 됩니다. 만약 완전 불변을 원하면 세터를 없애고, 필드가 모두 `final`인 `Builder`(필드 설정 후 `build()`가 불변 `ExcelCellStyleConfigurer`를 돌려주는 형태)로 대체할 수 있습니다.

- `ExcelCellStyleConfigurer`의 기본값은 `NoExcelColor`/`NoExcelAlign`/`NoExcelBorder`라 null이 들어가도 동작이 중단되지 않는 no-op 패턴을 제공합니다. 이를 통해 런타임 NPE를 차단하고, 실제 스타일을 지정한 경우에만 해당 부분이 적용되도록 의도한 것입니다.

- 스타일 적용 순서가 고정(색상 → 정렬 → 테두리)이라면 테스트나 문서에서 순서를 명시적으로 설명해주는 것이 안전합니다. 더 유연하게 만들려면 `List<Consumer<CellStyle>>`를 유지하다가 `apply` 시 순서대로 실행하는 작은 파이프라인으로 바꿀 수 있고, 이 리스트의 순서를 바꾸거나 추가/삭제하여 적용 흐름을 커스터마이징할 수 있습니다.

- `ColorPalette#fromInt`는 일부 팔레트 인덱스가 정의되어 있지 않아 매핑이 비어 있으면 예외를 던집니다. 전 인덱스를 채우거나, 현재처럼 누락된 인덱스를 만났을 때는 “지원되지 않는 팔레트 인덱스”임을 명시적으로 알려주는 메시지를 제공해 역직렬화 시 원인을 더 쉽게 파악할 수 있습니다.
