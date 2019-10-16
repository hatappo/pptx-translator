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

(def ts (time/format "yyyyMMddHHmmss" (time/local-date-time)))
(def region "ap-northeast-1")
(def srclang "ko") ; possible values @see https://docs.aws.amazon.com/translate/latest/dg/how-it-works.html
(def tgtlang "ja") ; possible values @see https://docs.aws.amazon.com/translate/latest/dg/how-it-works.html
(def infile "./resources/test.pptx")
(def outfile (str infile ".translated-at-" ts ".pptx"))

(with-open [ppt (XMLSlideShow. (io/input-stream infile))]
  (doseq [slide (.getSlides ppt)]
    (doseq [shape (.getShapes slide)
            :when (instance? XSLFTextShape shape)
            :let [text (.getText shape)
                  newtext (translate (str text))]]
      (-> shape .getTextBody (.setText newtext))))
  (with-open [dst (io/output-stream outfile)]
    (.write ppt dst)))

(defn- translate [text]
  ; (do (println text))
  (if-not (do (println text) (= 0 (-> text str clojure.string/trim .length)))
    (let [awsCreds (DefaultAWSCredentialsProviderChain/getInstance)
          b (AmazonTranslateClient/builder)
          tr (-> b (.withCredentials (AWSStaticCredentialsProvider. (.getCredentials awsCreds))) (.withRegion region) .build)
          req (-> (TranslateTextRequest.) (.withText (str text)) (.withSourceLanguageCode srclang) (.withTargetLanguageCode tgtlang))
          result (.translateText tr req)
          translated-text (.getTranslatedText result)]
      translated-text)))

(defn- translate-mock [text]
  (str "☆☆☆" text "★★★"))
