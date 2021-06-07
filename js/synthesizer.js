var resultsDiv, eventsDiv;
var highlightDiv;
var startSynthesisAsyncButton, pauseButton, resumeButton;
var updateVoiceListButton;

// subscription key and region for speech services.
var subscriptionKey, regionOptions;
var authorizationToken;
var voiceOptions, isSsml;
var SpeechSDK;
var synthesisText;
var synthesizer;
var player;
var wordBoundaryList = [];

document.addEventListener("DOMContentLoaded", function () {
    startSynthesisAsyncButton = document.getElementById("startSynthesisAsyncButton");
    updateVoiceListButton = document.getElementById("updateVoiceListButton");
    pauseButton = document.getElementById("pauseButton");
    resumeButton = document.getElementById("resumeButton");
    subscriptionKey = "6e4a67f034a545e0833179044dd6d172";
    regionOptions = "westeurope"
    resultsDiv = document.getElementById("resultsDiv");
    eventsDiv = document.getElementById("eventsDiv");
    voiceOptions = document.getElementById("voices");
    isSsml = false;
    highlightDiv = document.getElementById("highlightDiv");

    setInterval(function () {
        if (player !== undefined) {
            const currentTime = player.currentTime;
            var wordBoundary;
            for (const e of wordBoundaryList) {
                if (currentTime * 1000 > e.audioOffset / 10000) {
                    wordBoundary = e;
                } else {
                    break;
                }
            }
            if (wordBoundary !== undefined) {
                highlightDiv.innerHTML = synthesisText.value.substr(0, wordBoundary.textOffset) +
                    "<span class='highlight'>" + wordBoundary.text + "</span>" +
                    synthesisText.value.substr(wordBoundary.textOffset + wordBoundary.wordLength);
            } else {
                highlightDiv.innerHTML = synthesisText.value;
            }
        }
    }, 50);

    updateVoiceListButton.addEventListener("click", function () {
        var request = new XMLHttpRequest();
        request.open('GET',
            'https://' + regionOptions.value + ".tts.speech." +
            (regionOptions.value.startsWith("china") ? "azure.cn" : "microsoft.com") +
            "/cognitiveservices/voices/list", true);
        if (authorizationToken) {
            request.setRequestHeader("Authorization", "Bearer " + authorizationToken);
        } else {
            if (subscriptionKey.value === "" || subscriptionKey.value === "subscription") {
                alert("Please enter your Microsoft Cognitive Services Speech subscription key!");
                return;
            }
            request.setRequestHeader("Ocp-Apim-Subscription-Key", subscriptionKey.value);
        }

        request.onload = function () {
            if (request.status >= 200 && request.status < 400) {
                const response = this.response;
                const neuralSupport = (response.indexOf("AriaNeural") > 0);
                const defaultVoice = neuralSupport ? "AriaNeural" : "AriaRUS";
                let selectId;
                const data = JSON.parse(response);
                voiceOptions.innerHTML = "";
                data.forEach((voice, index) => {
                    voiceOptions.innerHTML += "<option value=\"" + voice.Name + "\">" + voice.Name + "</option>";
                    if (voice.Name.indexOf(defaultVoice) > 0) {
                        selectId = index;
                    }
                });
                voiceOptions.selectedIndex = selectId;
                voiceOptions.disabled = false;
            } else {
                window.console.log(this);
                eventsDiv.innerHTML += "cannot get voice list, code: " + this.status + " detail: " + this.statusText + "\r\n";
            }
        };

        request.send()
    });

    pauseButton.addEventListener("click", function () {
        player.pause();
        pauseButton.disabled = true;
        resumeButton.disabled = false;
    });

    resumeButton.addEventListener("click", function () {
        player.resume();
        pauseButton.disabled = false;
        resumeButton.disabled = true;
    });

    startSynthesisAsyncButton.addEventListener("click", function () {
        startSynthesisAsyncButton.disabled = true;
        resultsDiv.innerHTML = "";
        eventsDiv.innerHTML = "";
        wordBoundaryList = [];
        synthesisText = document.getElementById("synthesisText");

        // if we got an authorization token, use the token. Otherwise use the provided subscription key
        var speechConfig;
        if (authorizationToken) {
            speechConfig = SpeechSDK.SpeechConfig.fromAuthorizationToken(authorizationToken, regionOptions.value);
        } else {
            if (subscriptionKey.value === "" || subscriptionKey.value === "subscription") {
                alert("Please enter your Microsoft Cognitive Services Speech subscription key!");
                return;
            }
            speechConfig = SpeechSDK.SpeechConfig.fromSubscription(subscriptionKey.value, regionOptions.value);
        }

        speechConfig.speechSynthesisVoiceName = voiceOptions.value;

        player = new SpeechSDK.SpeakerAudioDestination();
        player.onAudioEnd = function (_) {
            window.console.log("playback finished");
            eventsDiv.innerHTML += "playback finished" + "\r\n";
            startSynthesisAsyncButton.disabled = false;
            pauseButton.disabled = true;
            resumeButton.disabled = true;
            wordBoundaryList = [];
        };

        var audioConfig = SpeechSDK.AudioConfig.fromSpeakerOutput(player);

        synthesizer = new SpeechSDK.SpeechSynthesizer(speechConfig, audioConfig);

        // The event synthesizing signals that a synthesized audio chunk is received.
        // You will receive one or more synthesizing events as a speech phrase is synthesized.
        // You can use this callback to streaming receive the synthesized audio.
        synthesizer.synthesizing = function (s, e) {
            window.console.log(e);
            eventsDiv.innerHTML += "(synthesizing) Reason: " + SpeechSDK.ResultReason[e.result.reason] +
                "Audio chunk length: " + e.result.audioData.byteLength + "\r\n";
        };

        // The synthesis started event signals that the synthesis is started.
        synthesizer.synthesisStarted = function (s, e) {
            window.console.log(e);
            eventsDiv.innerHTML += "(synthesis started)" + "\r\n";
            pauseButton.disabled = false;
        };

        // The event synthesis completed signals that the synthesis is completed.
        synthesizer.synthesisCompleted = function (s, e) {
            console.log(e);
            eventsDiv.innerHTML += "(synthesized)  Reason: " + SpeechSDK.ResultReason[e.result.reason] +
                " Audio length: " + e.result.audioData.byteLength + "\r\n";
        };

        // The event signals that the service has stopped processing speech.
        // This can happen when an error is encountered.
        synthesizer.SynthesisCanceled = function (s, e) {
            const cancellationDetails = SpeechSDK.CancellationDetails.fromResult(e.result);
            let str = "(cancel) Reason: " + SpeechSDK.CancellationReason[cancellationDetails.reason];
            if (cancellationDetails.reason === SpeechSDK.CancellationReason.Error) {
                str += ": " + e.result.errorDetails;
            }
            window.console.log(e);
            eventsDiv.innerHTML += str + "\r\n";
            startSynthesisAsyncButton.disabled = false;
            pauseButton.disabled = true;
            resumeButton.disabled = true;
        };

        // This event signals that word boundary is received. This indicates the audio boundary of each word.
        // The unit of e.audioOffset is tick (1 tick = 100 nanoseconds), divide by 10,000 to convert to milliseconds.
        synthesizer.wordBoundary = function (s, e) {
            window.console.log(e);
            eventsDiv.innerHTML += "(WordBoundary), Text: " + e.text + ", Audio offset: " + e.audioOffset / 10000 + "ms." + "\r\n";
            wordBoundaryList.push(e);
        };

        const complete_cb = function (result) {
            if (result.reason === SpeechSDK.ResultReason.SynthesizingAudioCompleted) {
                resultsDiv.innerHTML += "synthesis finished";
            } else if (result.reason === SpeechSDK.ResultReason.Canceled) {
                resultsDiv.innerHTML += "synthesis failed. Error detail: " + result.errorDetails;
            }
            window.console.log(result);
            synthesizer.close();
            synthesizer = undefined;
        };
        const err_cb = function (err) {
            startSynthesisAsyncButton.disabled = false;
            phraseDiv.innerHTML += err;
            window.console.log(err);
            synthesizer.close();
            synthesizer = undefined;
        };
        if (isSsml.checked) {
            synthesizer.speakSsmlAsync(synthesisText.value,
                complete_cb,
                err_cb);
        } else {
            synthesizer.speakTextAsync(synthesisText.value,
                complete_cb,
                err_cb);
        }
    });

    Initialize(function (speechSdk) {
        SpeechSDK = speechSdk;
        startSynthesisAsyncButton.disabled = false;
        pauseButton.disabled = true;
        resumeButton.disabled = true;

        // in case we have a function for getting an authorization token, call it.
        if (typeof RequestAuthorizationToken === "function") {
            RequestAuthorizationToken();
        }
    });
});