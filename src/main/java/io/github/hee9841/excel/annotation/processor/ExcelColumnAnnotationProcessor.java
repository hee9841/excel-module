package io.github.hee9841.excel.annotation.processor;

import static io.github.hee9841.excel.global.SystemValues.ALLOWED_FIELD_TYPES;
import static io.github.hee9841.excel.global.SystemValues.ALLOWED_FIELD_TYPES_STRING;

import io.github.hee9841.excel.annotation.ExcelColumn;
import java.util.Objects;
import java.util.Set;
import javax.annotation.processing.AbstractProcessor;
import javax.annotation.processing.Messager;
import javax.annotation.processing.ProcessingEnvironment;
import javax.annotation.processing.Processor;
import javax.annotation.processing.RoundEnvironment;
import javax.annotation.processing.SupportedAnnotationTypes;
import javax.annotation.processing.SupportedSourceVersion;
import javax.lang.model.SourceVersion;
import javax.lang.model.element.Element;
import javax.lang.model.element.ElementKind;
import javax.lang.model.element.TypeElement;
import javax.lang.model.type.TypeKind;
import javax.lang.model.type.TypeMirror;
import javax.lang.model.util.Elements;
import javax.lang.model.util.Types;
import javax.tools.Diagnostic;

/**
 * Annotation processor for {@link ExcelColumn} annotation.
 * This processor validates that the annotation is only applied to supported field types.
 *
 * <p>Supported types include:
 * <ul>
 *   <li>String</li>
 *   <li>Character/char</li>
 *   <li>Numeric types (Byte, Short, Integer, Long, Float, Double and their primitives)</li>
 *   <li>Boolean/boolean</li>
 *   <li>Date/Time types (LocalDate, LocalDateTime, Date, java.sql.Date)</li>
 *   <li>Enum types</li>
 * </ul>
 */
@SupportedAnnotationTypes("io.github.hee9841.excel.annotation.ExcelColumn")
@SupportedSourceVersion(SourceVersion.RELEASE_8)
public class ExcelColumnAnnotationProcessor extends AbstractProcessor {

    private Messager messager;
    private Types typeUtils;
    private Elements elementUtils;

    @Override
    public synchronized void init(ProcessingEnvironment processingEnv) {
        super.init(processingEnv);
        messager = processingEnv.getMessager();
        typeUtils = processingEnv.getTypeUtils();
        elementUtils = processingEnv.getElementUtils();
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        boolean hasError = false;

        for (Element element : roundEnv.getElementsAnnotatedWith(ExcelColumn.class)) {
            if (!isAllowedType(element)) {
                hasError = true;
            }
        }
        return !hasError;
    }

    /**
     * Checks if the annotated element has an allowed type.
     * Validates that the element is a field and checks if its type is either an enum
     * or one of the supported types defined in ALLOWED_TYPES. Array types are not allowed.
     *
     * @param element the element to check
     * @return true if the element has an allowed type, false otherwise
     */
    private boolean isAllowedType(Element element) {
        // 1.If element is enum type, return true
        if (!element.getKind().isField()) {
            error(element, "@ExcelColumn can only be applied to field type");
            return false;
        }

        TypeMirror typeMirror = element.asType();

        // 3. If element is array type, return false
        if (typeMirror.getKind() == TypeKind.ARRAY) {
            error(element, "@ExcelColumn cannot be applied to array type");
            return false;
        }

        if (isEnumType(typeMirror)) {
            return true;
        }


        // 4.Primitive type check
        if (typeMirror.getKind().isPrimitive()) {
            return true;
        }

        //5.Check if type is in allowed types list
        if (typeMirror.getKind() == TypeKind.DECLARED) {
            boolean isAllowed = ALLOWED_FIELD_TYPES.stream()
                .map(clazz -> elementUtils.getTypeElement(clazz.getTypeName()))
                .filter(Objects::nonNull)
                .map(TypeElement::asType)
                .anyMatch(allowedType ->
                    typeUtils.isAssignable(typeMirror, allowedType));

            if (!isAllowed) {
                error(element, "@ExcelColumn can only be applied to allowed types(%s).",
                    ALLOWED_FIELD_TYPES_STRING);
                return false;
            }

            return true;
        }

        return false;
    }

    /**
     * Checks if the given TypeMirror represents an enum type.
     *
     * @param typeMirror the type to check
     * @return true if the type is an enum, false otherwise
     */
    private boolean isEnumType(TypeMirror typeMirror) {
        TypeElement typeElement = (TypeElement) typeUtils.asElement(typeMirror);
        return typeElement != null && typeElement.getKind() == ElementKind.ENUM;
    }

    /**
     * Reports an error for the given element using the processor's messager.
     *
     * @param e the element for which to report the error
     * @param msg the error message format string
     * @param args the arguments to be used in the formatted message
     */
    private void error(Element e, String msg, Object... args) {
        messager.printMessage(
            Diagnostic.Kind.ERROR,
            String.format(msg, args),
            e);
    }
}
