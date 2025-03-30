package io.github.hee9841.excel.annotation.processor;

import io.github.hee9841.excel.annotation.Excel;
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
import javax.tools.Diagnostic;

/**
 * Annotation processor for handling {@code @Excel} annotations.
 * This processor validates classes annotated with {@code @Excel} to ensure they meet
 * the required criteria for Excel processing.
 *
 * <p>Supported criteria:
 * <ul>
 *     <li>Can be applied to regular classes or record classes</li>
 *     <li>Target class must not be abstract</li>
 * </ul>
 */
@SupportedAnnotationTypes("io.github.hee9841.excel.annotation.Excel")
public class ExcelAnnotationProcessor extends AbstractProcessor {

    private Messager messager;

    @Override
    public synchronized void init(ProcessingEnvironment processingEnv) {
        super.init(processingEnv);
        messager = processingEnv.getMessager();
    }

    @Override
    public SourceVersion getSupportedSourceVersion() {
        return SourceVersion.latestSupported();
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        boolean hasError = false;

        for (Element element : roundEnv.getElementsAnnotatedWith(Excel.class)) {
            if (!isValidExcelClass(element)) {
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
     * Reports an error for the given element.
     *
     * @param e    the element for which to report the error
     * @param msg  the error message with format specifiers
     * @param args arguments referenced by the format specifiers in the message
     */
    private void error(Element e, String msg, Object... args) {
        messager.printMessage(
            Diagnostic.Kind.ERROR,
            String.format(msg, args),
            e);
    }
}
