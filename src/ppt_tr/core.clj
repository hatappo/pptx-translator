(ns ppt-tr.core
  (:require
   [clojure.java.io :as io]
   [java-time :as time])
  (:import 
   (com.amazonaws.auth AWSStaticCredentialsProvider BasicAWSCredentials DefaultAWSCredentialsProviderChain)
   (com.amazonaws.client.builder AwsClientBuilder)
   (com.amazonaws.services.translate AmazonTranslate AmazonTranslateClient)
   (com.amazonaws.services.translate.model TranslateTextRequest TranslateTextResult)
   (org.apache.poi.xslf.usermodel XMLSlideShow XSLFShape XSLFTextShape XSLFSlide)))

(def region "ap-northeast-1")

; possible values as lang-from lang-to @see https://docs.aws.amazon.com/translate/latest/dg/how-it-works.html
(defn- translate [text lang-from lang-to]
  (if-not (do (println text) (= 0 (-> text str clojure.string/trim .length)))
    (let [awsCreds (DefaultAWSCredentialsProviderChain/getInstance)
          b (AmazonTranslateClient/builder)
          tr (-> b (.withCredentials (AWSStaticCredentialsProvider. (.getCredentials awsCreds))) (.withRegion region) .build)
          req (-> (TranslateTextRequest.) (.withText (str text)) (.withSourceLanguageCode lang-from) (.withTargetLanguageCode lang-to))
          result (.translateText tr req)
          translated-text (.getTranslatedText result)]
      translated-text)))

(defn- translate-pptx [org dst lang-from lang-to]
  (with-open [ppt (XMLSlideShow. (io/input-stream org))]
    (doseq [slide (.getSlides ppt)]
      (doseq [shape (.getShapes slide)
              :when (instance? XSLFTextShape shape)
              :let [text (.getText shape)
                    newtext (translate (str text) lang-from lang-to)]]
        (-> shape .getTextBody (.setText newtext))))
    (with-open [dst (io/output-stream dst)]
      (.write ppt dst))))

(defn -main []
  (let [ts (time/format "yyyyMMddHHmmss" (time/local-date-time))
        org "./resources/test.pptx"
        dst (str org ".translated-at-" ts ".pptx")
        from "ko"
        to "ja"]
    (translate-pptx org dst from to)))

