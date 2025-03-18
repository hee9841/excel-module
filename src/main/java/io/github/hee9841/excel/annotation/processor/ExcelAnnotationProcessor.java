package io.github.hee9841.excel.annotation.processor;

import com.google.auto.service.AutoService;
import io.github.hee9841.excel.annotation.Excel;
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
import javax.lang.model.element.Modifier;
import javax.lang.model.element.TypeElement;
import javax.tools.Diagnostic;


@AutoService(Processor.class)
@SupportedAnnotationTypes("io.github.hee9841.excel.annotation.Excel")
@SupportedSourceVersion(SourceVersion.RELEASE_8)
public class ExcelAnnotationProcessor extends AbstractProcessor {

    private Messager messager;

    @Override
    public synchronized void init(ProcessingEnvironment processingEnv) {
        super.init(processingEnv);
        messager = processingEnv.getMessager();
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

    private boolean isValidExcelClass(Element element) {
        // Java 14 이상에서 record 클래스 체크
        if (element.getKind().name().equals("RECORD")) {
            return true;
        }
        // class 가 이니면
        if (element.getKind() != ElementKind.CLASS) {
            error(element,
                "@%s can only be applied to classes or record classes",
                Excel.class.getSimpleName()
            );
            return false;
        }

        TypeElement typeElement = (TypeElement) element;
        // 클래스는 추상 클래스가 아니어야한다.
        if (typeElement.getModifiers().contains(Modifier.ABSTRACT)) {
            error(element,
                "The class %s is abstract. You can't annotate abstract classes with @%s",
                typeElement.getQualifiedName().toString(), Excel.class.getSimpleName());

            return false;
        }

        return true;
    }

    private void error(Element e, String msg, Object... args) {
        messager.printMessage(
            Diagnostic.Kind.ERROR,
            String.format(msg, args),
            e);
    }
}
