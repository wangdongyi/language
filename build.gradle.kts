import org.gradle.api.file.DuplicatesStrategy
plugins {
    id("java")
    id("org.jetbrains.kotlin.jvm") version "1.9.25"
    id("org.jetbrains.intellij") version "1.17.4"
}

group = "com.wdy.language"
version = "1.5.0-WDY"

repositories {
    mavenCentral()
}
dependencies {
    implementation(kotlin("stdlib"))
    implementation("org.apache.poi:poi:5.2.3")
    implementation("org.apache.poi:poi-ooxml:5.2.3")
    implementation("org.apache.xmlbeans:xmlbeans:5.1.1")
    implementation("org.apache.commons:commons-collections4:4.4")
    implementation("org.apache.poi:poi-ooxml-schemas:4.1.2")
    implementation("org.apache.xmlbeans:xmlbeans:5.1.1")
    implementation("org.apache.commons:commons-compress:1.21")
    implementation("com.fasterxml.woodstox:woodstox-core:6.5.1") // 解决 XML 依赖冲突
}


sourceSets {
    main {
        resources {
            // 修复1：明确包含图片资源
            srcDir("src/main/resources")
            include("META-INF/**", "**/*.jpg") // 添加图片类型包含
        }
    }
}

// Configure Gradle IntelliJ Plugin
// Read more: https://plugins.jetbrains.com/docs/intellij/tools-gradle-intellij-plugin.html
intellij {
    version.set("2024.1.7")
    type.set("IC") // Target IDE Platform

    plugins.set(listOf(/* Plugin Dependencies */))
}

tasks {
    // Set the JVM compatibility versions
    withType<JavaCompile> {
        sourceCompatibility = "17"
        targetCompatibility = "17"
    }
    withType<org.jetbrains.kotlin.gradle.tasks.KotlinCompile> {
        kotlinOptions.jvmTarget = "17"
    }

    patchPluginXml {
        sinceBuild.set("193")
        untilBuild.set("259.*")
    }

    signPlugin {
        certificateChain.set(System.getenv("CERTIFICATE_CHAIN"))
        privateKey.set(System.getenv("PRIVATE_KEY"))
        password.set(System.getenv("PRIVATE_KEY_PASSWORD"))
    }

    publishPlugin {
        token.set(System.getenv("PUBLISH_TOKEN"))
    }
    processResources{
        duplicatesStrategy = DuplicatesStrategy.EXCLUDE
        // 修复3：确保包含所有资源文件
        from("src/main/resources") {
            include("META-INF/**")
            include("**/*.jpg") // 明确包含JPG文件
        }
    }
}
tasks.register<Copy>("copyImagesToPluginMeta") {
    from("src/main/resources/META-INF/images") // 图片原始位置
    into("$buildDir/distributions/META-INF/images") // 复制到插件根目录
    duplicatesStrategy = DuplicatesStrategy.EXCLUDE
}

// 确保构建插件前先复制图片
tasks.named("buildPlugin") {
    dependsOn("copyImagesToPluginMeta")
}
tasks.processResources {
    from("src/main/resources/META-INF") {
        into("META-INF")
    }
}
