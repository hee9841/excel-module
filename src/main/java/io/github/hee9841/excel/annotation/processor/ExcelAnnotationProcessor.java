package io.github.hee9841.excel.annotation.processor;

import static io.github.hee9841.excel.global.SystemValues.ALLOWED_FIELD_TYPES;
import static io.github.hee9841.excel.global.SystemValues.ALLOWED_FIELD_TYPES_STRING;

import io.github.hee9841.excel.annotation.Excel;
import io.github.hee9841.excel.annotation.ExcelColumn;
import java.util.Objects;
import java.util.Set;
import javax.annotation.processing.AbstractProcessor;
import javax.annotation.processing.Messager;
import javax.annotation.processing.ProcessingEnvironment;
import javax.annotation.processing.RoundEnvironment;
import javax.annotation.processing.SupportedAnnotationTypes;
import javax.lang.model.SourceVersion;
import javax.lang.model.element.Element;
import javax.lang.model.element.ElementKind;
import javax.lang.model.element.Modifier;
import javax.lang.model.element.TypeElement;
import javax.lang.model.type.TypeKind;
import javax.lang.model.type.TypeMirror;
import javax.lang.model.util.Elements;
import javax.lang.model.util.Types;
import javax.tools.Diagnostic;

/**
 * Annotation processor for handling {@code @Excel} and {@code @ExcelColumn} annotations.
 * This processor validates the relationship between these annotations and ensures they meet
 * the required criteria for Excel processing.
 *
 * <p>Validation rules:
 * <ul>
 *     <li>Classes annotated with {@code @Excel} must have at least one field annotated with {@code @ExcelColumn}</li>
 *     <li>Fields annotated with {@code @ExcelColumn} must be in a class annotated with {@code @Excel}</li>
 *     <li>Classes annotated with {@code @Excel} must be either regular classes or record classes</li>
 *     <li>Regular classes annotated with {@code @Excel} must not be abstract</li>
 *     <li>Fields annotated with {@code @ExcelColumn} must be of supported types</li>
 * </ul>
 *
 * <p>Supported field types for {@code @ExcelColumn}:
 * <ul>
 *   <li>String</li>
 *   <li>Character/char</li>
 *   <li>Numeric types (Byte, Short, Integer, Long, Float, Double and their primitives)</li>
 *   <li>Boolean/boolean</li>
 *   <li>Date/Time types (LocalDate, LocalDateTime, Date, java.sql.Date)</li>
 *   <li>Enum types</li>
 * </ul>
 *
 * <p>Note: Array types are not supported for {@code @ExcelColumn} fields.
 */
@SupportedAnnotationTypes({
    "io.github.hee9841.excel.annotation.Excel",
    "io.github.hee9841.excel.annotation.ExcelColumn"
})
public class ExcelAnnotationProcessor extends AbstractProcessor {

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
    public SourceVersion getSupportedSourceVersion() {
        return SourceVersion.latestSupported();
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        boolean hasError = false;


        // First pass: collect all Excel-annotated classes
        for (Element element : roundEnv.getElementsAnnotatedWith(Excel.class)) {
            if (!isValidExcelClass(element)) {
                hasError = true;
            }
        }

        // Second pass: validate ExcelColumn annotations
        for (Element element : roundEnv.getElementsAnnotatedWith(ExcelColumn.class)) {
            if (!isAllowedType(element)) {
                hasError = true;
            }
        }

        return !hasError;
    }

    /**
     * Validates if the element annotated with {@code @Excel} meets the required criteria.
     * The following conditions are checked:
     * <ul>
     *     <li>Element must be either a regular class or a record class</li>
     *     <li>If it's a regular class, it must not be abstract</li>
     * </ul>
     *
     * @param element the element to validate, typically a class or record annotated with
     *                {@code @Excel}
     * @return true if the element meets all validation criteria, false if any validation fails
     * (error messages will be reported via {@link #error(Element, String, Object...)})
     * @see Excel
     */
    private boolean isValidExcelClass(Element element) {

        // Check for record class in Java 14 and above
        if (element.getKind().name().equals("RECORD")) {
            return true;
        }

        if (element.getKind() != ElementKind.CLASS) {
            error(element,
                "@%s can only be applied to classes or record classes",
                Excel.class.getSimpleName()
            );
            return false;
        }

        TypeElement typeElement = (TypeElement) element;

        // The class must not be an abstract class.
        if (typeElement.getModifiers().contains(Modifier.ABSTRACT)) {
            error(element,
                "The class %s is abstract. You can't annotate abstract classes with @%s",
                typeElement.getQualifiedName().toString(), Excel.class.getSimpleName());

            return false;
        }

        return true;
    }


    /**
     * Checks if the annotated element has an allowed field type.
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
     * Reports an error for the given element using the processor's message.
     *
     * @param e    the element for which to report the error
     * @param msg  the error message format string
     * @param args the arguments to be used in the formatted message
     */
    private void error(Element e, String msg, Object... args) {
        messager.printMessage(
            Diagnostic.Kind.ERROR,
            String.format(msg, args),
            e);
    }
}
