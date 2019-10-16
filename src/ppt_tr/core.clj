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

; reference for POI
; https://www.tutorialspoint.com/apache_poi_ppt/apache_poi_ppt_formatting_text.htm
; https://qiita.com/Lain_/items/6a1d3ed0255720d6d57a
; https://gist.github.com/ponkore/4216377
; https://poi.apache.org/apidocs/dev/org/apache/poi/xslf/usermodel/XSLFTextBox.html
; reference for Amazon Translate and Java SDK
; https://docs.aws.amazon.com/translate/latest/dg/examples-java.html












; (with-open [org (->> infile  File. FileInputStream. XMLSlideShow.)]
;   (doseq [slide (.getSlides org)]
;     (doseq [shape (.getShapes slide)]
;       (-> shape .getTextBody (.setText "できたー！！！"))))
;   (def dst (->> outfile File. FileOutputStream. (.write org))))


; (def org (->> infile (new File) (new FileInputStream) (new XMLSlideShow)))




; (io/copy (io/file infile) (io/file dst))




; (io/copy (io/file org) (io/file dst))

; (def pptx (new XSLFSlideShow dst))

; (def list (-> pptx .getSlideReferences .getSldIdList)) 

; (doseq [entry list]
;   (let [slide (-> pptx (.getSlide entry))]
;     (println (class slide))
;     ))


; (for [slide list]
;   (let [shapes (-> pptx (.getSlide slide) .getCSld .getSpTree .getSpList)]
;     (for [shape shapes]
;       (prn shape))))

; (def slides (for [slide list] slide))

; (map #(-> pptx (.getSlide %)) list)
