/**
 * RAPID MARK - Local NLP Engine (Ultimate Hybrid Phase 1)
 * 100% Client-Side. No APIs. No Cloud Data Storage.
 * Designed specifically for Teacher-Facing Marking Dashboards.
 */

const RapidEngine = (function() {
    
    // ==========================================
    // 1. CALIBRATION THRESHOLDS
    // ==========================================
    
    // Centralized thresholds make tuning the engine incredibly easy.
    const CALIBRATION = {
        upper_primary: {
            runOnWordLimit: 20, runOnPunctMin: 1, stopWordMax: 2,
            vagueTiers: 1, repetitionStemLength: 4, tenseWindow: 2
        },
        lower_sec: {
            runOnWordLimit: 30, runOnPunctMin: 2, stopWordMax: 1,
            vagueTiers: 2, repetitionStemLength: 4, tenseWindow: 3
        },
        upper_sec: {
            runOnWordLimit: 40, runOnPunctMin: 3, stopWordMax: 0,
            vagueTiers: 3, repetitionStemLength: 5, tenseWindow: 3
        }
    };

    // ==========================================
    // 2. DICTIONARIES
    // ==========================================
    
    const DICTS = {
        fanboys: ['and', 'but', 'so', 'or', 'yet', 'for', 'nor'],
        subordinators: ['because', 'although', 'even though', 'since', 'unless', 'whereas', 'while', 'which', 'that', 'if', 'until', 'before', 'after'],
        
        vagueTier1: ['good', 'bad', 'stuff', 'things', 'a lot of', 'huge'],
        vagueTier2: ['many things', 'huge impact', "in today's society", 'back in the day', 'changed everything', 'played a role'],
        vagueTier3: ['since the dawn of time', 'throughout history', 'human nature', 'society says', 'it is clear that', 'undeniable', 'needless to say'],
        
        stopWords: ['really', 'very', 'basically', 'literally', 'actually', 'just', 'obviously', 'kind of', 'sort of', 'totally', 'essentially', 'honestly'],
        reportingVerbs: ['said', 'stated', 'argued', 'writes', 'notes', 'claims', 'suggests', 'implies', 'demonstrates', 'shows', 'explains', 'reveals', 'highlights', 'asserts', 'contends'],
        
        pastMarkers: ['was', 'were', 'had', 'did', 'argued', 'showed', 'revealed', 'described', 'explained', 'fought', 'went', 'made', 'saw', 'told', 'came', 'became', 'looked', 'seemed', 'tried', 'wanted', 'needed', 'used'],
        presentMarkers: ['is', 'are', 'has', 'does', 'argues', 'shows', 'reveals', 'describes', 'explains', 'fights', 'goes', 'makes', 'sees', 'tells', 'comes', 'becomes', 'looks', 'seems', 'tries', 'wants', 'needs', 'uses'],
        
        explanationStarters: ['this shows', 'this means', 'this highlights', 'this suggests', 'this reveals', 'this demonstrates', 'this tells the reader'],
        explanationWeakEndings: ['is important', 'is sad', 'is bad', 'is good', 'is powerful', 'is effective', 'a big impact', 'an impact', 'the theme', 'the issue', 'the problem'],
        explanationDepthMarkers: ['because', 'therefore', 'thereby', 'suggesting', 'revealing', 'implying', 'indicating', 'emphasising', 'highlighting', 'which', 'thus', 'hence'],
        techniqueWords: ['metaphor', 'imagery', 'symbolism', 'tone', 'contrast', 'juxtaposition', 'alliteration', 'repetition', 'characterisation', 'setting', 'dialogue', 'language', 'structure'],
        
        repeatedIdeaIgnoreWords: ['the', 'a', 'an', 'it', 'this', 'that', 'they', 'he', 'she', 'we', 'to', 'of', 'and', 'in', 'on', 'for', 'with', 'at', 'by', 'from', 'is', 'was', 'are', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'as', 'but', 'or', 'if', 'because', 'then', 'so', 'than', 'into', 'about', 'over', 'after', 'before'],

        abbreviations: ['mr', 'mrs', 'ms', 'dr', 'prof', 'st', 'vs', 'i.e', 'e.g', 'etc']
    };

    // ==========================================
    // 3. PRE-PROCESSING & TOKENIZATION
    // ==========================================

    function splitIntoSentences(text) {
        let safeText = text.replace(/[“”]/g, '"').replace(/[‘’]/g, "'").replace(/\.\.\./g, "___ELLIPSIS___");
        
        DICTS.abbreviations.forEach(abbr => {
            const regex = new RegExp(`\\b${abbr}\\.`, 'gi');
            safeText = safeText.replace(regex, `${abbr}@@@`);
        });

        const sentenceRegex = /([.!?]+)(?=\s+["']?[A-Za-z])/g;
        const splitTokens = safeText.replace(sentenceRegex, "$1|SPLIT|");
        let rawSentences = splitTokens.split("|SPLIT|");

        return rawSentences.map((raw, index) => {
            let restored = raw.replace(/@@@/g, '.').replace(/___ELLIPSIS___/g, "...");
            let trimmed = restored.trim();
            if (!trimmed) return null;

            const words = trimmed.toLowerCase().replace(/[^a-z0-9'\s-]/gi, '').split(/\s+/).filter(Boolean);
            
            return {
                id: index,
                raw: restored,
                trimmed: trimmed,
                words: words,
                wordCount: words.length,
                commaCount: (trimmed.match(/,/g) || []).length,
                conjunctionCount: words.filter(w => DICTS.fanboys.includes(w)).length,
                hasQuote: /["']/.test(trimmed),
                startsWithQuote: /^\s*["']/.test(trimmed)
            };
        }).filter(Boolean);
    }

    // Claude's Lightweight Stemmer - Massive accuracy boost for Repetition logic
    function stem(word) {
        let w = word.toLowerCase();
        const rules = [
            [/ies$/, "y"], [/ied$/, "y"], [/ying$/, "y"], [/ves$/, "f"],
            [/ation$/, "ate"], [/tions?$/, "te"], [/ings?$/, ""],
            [/edly$/, "e"], [/ness$/, ""], [/ment$/, ""], [/ful$/, ""],
            [/less$/, ""], [/ity$/, ""], [/ly$/, ""], [/er$/, ""],
            [/est$/, ""], [/ed$/, ""], [/es$/, ""], [/s$/, ""]
        ];
        for (const [pattern, replacement] of rules) {
            if (pattern.test(w)) {
                const stemmed = w.replace(pattern, replacement);
                if (stemmed.length >= 3) return stemmed;
            }
        }
        return w;
    }

    function getContentStems(words) {
        return words
            .filter(w => !DICTS.repeatedIdeaIgnoreWords.includes(w))
            .map(stem)
            .filter(w => w.length > 2);
    }

    // ==========================================
    // 4. CORE DETECTORS
    // ==========================================

    function detectRunOn(sentence, config) {
        const wCount = sentence.wordCount, cCount = sentence.commaCount, conjCount = sentence.conjunctionCount;
        let flagged = false;

        if (wCount > config.runOnWordLimit && cCount < config.runOnPunctMin) flagged = true;
        // Polysyndeton check (too many 'ands')
        if (conjCount >= 4) flagged = true;

        if (flagged) return createFlag("SPLIT THIS RUN-ON", sentence, "Exceeds length threshold with insufficient punctuation.", { wordCount: wCount, commaCount: cCount, conjunctions: conjCount });
        return null;
    }

    function detectFragment(sentence) {
        if (sentence.wordCount === 0) return null;
        const firstWord = sentence.words[0];
        
        if (DICTS.fanboys.includes(firstWord) && sentence.wordCount < 8) {
            return createFlag("FIX THIS FRAGMENT", sentence, "Starts with coordinating conjunction.", { firstWord, wordCount: sentence.wordCount });
        }
        if (DICTS.subordinators.includes(firstWord) && sentence.commaCount === 0) {
            return createFlag("FIX THIS FRAGMENT", sentence, "Subordinate clause missing main clause.", { firstWord, commaCount: sentence.commaCount });
        }
        return null;
    }

    function detectVagueStatement(sentence, config) {
        let activeLists = [...DICTS.vagueTier1];
        if (config.vagueTiers >= 2) activeLists.push(...DICTS.vagueTier2);
        if (config.vagueTiers === 3) activeLists.push(...DICTS.vagueTier3);

        const textLower = sentence.trimmed.toLowerCase();
        
        for (const phrase of activeLists) {
            if (new RegExp(`\\b${phrase}\\b`, 'i').test(textLower)) {
                return createFlag("VAGUE STATEMENT", sentence, "Imprecise phrasing detected.", { matchedPhrase: phrase });
            }
        }

        // Claude's Nominalisation Check (Upper Secondary Only)
        if (config.vagueTiers === 3) {
            const nominalMatch = sentence.trimmed.match(/\b(the \w+tion of|the \w+ment of|the \w+ance of)\b/i);
            if (nominalMatch) {
                return createFlag("VAGUE STATEMENT", sentence, "Wordy nominalisation detected.", { matchedPhrase: nominalMatch[0] });
            }
        }
        return null;
    }

    function detectStopWords(sentence, config) {
        let matchedWords = sentence.words.filter(w => DICTS.stopWords.includes(w));
        
        if (matchedWords.length > config.stopWordMax) {
            return createFlag("TOO MANY STOP WORDS", sentence, "Overuse of conversational filler.", { count: matchedWords.length, words: matchedWords });
        }
        return null;
    }

    function detectUnintroducedQuote(sentence, config) {
        if (!sentence.hasQuote) return null;
        const quoteIndex = (sentence.trimmed.match(/["']/) || {}).index || -1;
        
        if (sentence.startsWithQuote) {
            return createFlag("QUOTE NOT INTRODUCED", sentence, "Dropped quote.", { startsWithQuote: true });
        }
        
        if (quoteIndex > 0) {
            const wordsBefore = sentence.trimmed.substring(0, quoteIndex).toLowerCase().replace(/[^a-z\s]/g, '').split(/\s+/).filter(Boolean);
            if (!wordsBefore.slice(-4).some(w => DICTS.reportingVerbs.includes(w))) {
                return createFlag("QUOTE NOT INTRODUCED", sentence, "Missing analytical reporting verb.", { wordsBeforeQuote: wordsBefore.slice(-4) });
            }
        }
        return null;
    }

    function detectMixingTense(sentence, index, sentences, config) {
        // Claude's Sliding Window Logic
        const windowStart = Math.max(0, index - Math.floor(config.tenseWindow / 2));
        const windowEnd = Math.min(sentences.length, index + Math.ceil(config.tenseWindow / 2));
        const contextSentences = sentences.slice(windowStart, windowEnd);

        let totalPast = 0, totalPresent = 0;
        let pastWords = [], presentWords = [];

        contextSentences.forEach(s => {
            const past = s.words.filter(w => DICTS.pastMarkers.includes(w));
            const present = s.words.filter(w => DICTS.presentMarkers.includes(w));
            totalPast += past.length;
            totalPresent += present.length;
            pastWords.push(...past);
            presentWords.push(...present);
        });

        // If strong collision in the local window (excluding quotes)
        if (totalPast >= 2 && totalPresent >= 2 && !sentence.hasQuote) {
            return createFlag("MIXING TENSE", sentence, "Tense collision in surrounding sentences.", { 
                pastVerbs: [...new Set(pastWords)], 
                presentVerbs: [...new Set(presentWords)] 
            });
        }
        return null;
    }

    function detectRepetition(sentence, index, sentences, config) {
        if (index === 0) return null;
        const prevSentence = sentences[index - 1];
        
        const currentStems = getContentStems(sentence.words).filter(s => s.length >= config.repetitionStemLength);
        const prevStems = getContentStems(prevSentence.words).filter(s => s.length >= config.repetitionStemLength);
        
        const repeatedStems = currentStems.filter(stem => prevStems.includes(stem));
        const uniqueRepeats = [...new Set(repeatedStems)];
        
        if (uniqueRepeats.length >= 2) {
            return createFlag("REPEATED IDEA / REPETITION", sentence, "Overlapping root vocabulary with previous sentence.", { repeatedStems: uniqueRepeats });
        }
        return null;
    }

    function detectExplainEvidence(sentence) {
        const textLower = sentence.trimmed.toLowerCase();
        const startsWithFrame = DICTS.explanationStarters.some(starter => textLower.startsWith(starter));
        const hasWeakEnding = DICTS.explanationWeakEndings.some(ending => textLower.includes(ending));
        const lacksDepth = !DICTS.explanationDepthMarkers.some(marker => textLower.includes(marker));

        if (startsWithFrame && lacksDepth && hasWeakEnding) {
            return createFlag("EXPLAIN THE EVIDENCE / ADD MORE DEPTH", sentence, "Analytical bottleneck. Stays surface-level.", { startsWithFrame: true, lacksDepth: true });
        }
        return null;
    }

    function detectProofreadingNeeded(sentence) {
        const text = sentence.raw;
        let issues = [];

        if (/^\s*[a-z]/.test(text)) issues.push("Starts with lowercase letter");
        if (/(^|[\s(])i([\s,.!?;:)\]])/g.test(text)) issues.push("Lowercase 'i'");
        if (/\s+[,.!?;:]/.test(text)) issues.push("Space before punctuation");
        if (/ {2,}/.test(text)) issues.push("Double spacing");

        if (issues.length > 0) {
            return createFlag("PROOFREADING NEEDED", sentence, "Visible mechanical errors.", { proofreadingIssues: issues });
        }
        return null;
    }

    // ==========================================
    // 5. UTILITIES & API
    // ==========================================

    function createFlag(code, sentence, reason, evidence) {
        return {
            code,
            sentenceIndex: sentence.id,
            spanText: sentence.trimmed,
            reason,
            evidence
        };
    }

    return {
        analyze: function(text, level = 'lower_sec') {
            const config = CALIBRATION[level] || CALIBRATION['lower_sec'];
            const sentences = splitIntoSentences(text);
            const flags = [];

            sentences.forEach((sentence, index) => {
                const results = [
                    detectRunOn(sentence, config),
                    detectFragment(sentence),
                    detectVagueStatement(sentence, config),
                    detectStopWords(sentence, config),
                    detectUnintroducedQuote(sentence, config),
                    detectMixingTense(sentence, index, sentences, config),
                    detectRepetition(sentence, index, sentences, config),
                    detectExplainEvidence(sentence),
                    detectProofreadingNeeded(sentence)
                ].filter(Boolean);

                flags.push(...results);
            });

            const seen = new Set();
            return flags.filter(f => {
                const key = `${f.code}::${f.sentenceIndex}`;
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
        }
    };
})();

/* ==========================================
   QUIKDRAFT SEND MODULE
   ========================================== */


const QuikDraftSend = (function() {
    const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbw1DuwokbFs5Dl-v21gTSjxWJ6Z11eQpYgD6f5VxFp6oMKlfUdM_6-COxKkEdHhLbb-uQ/exec';

    async function sendFeedback(row, feedbackText, feedbackHtml) {
        if (!row || Number(row) < 2) {
            throw new Error('Missing valid sheet row.');
        }

        const text = feedbackText != null ? String(feedbackText) : '';
        const html = feedbackHtml != null ? String(feedbackHtml) : '';

        if (!text.trim() && !html.trim()) {
            throw new Error('Missing feedback text.');
        }

        const response = await fetch(WEB_APP_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain;charset=utf-8'
            },
            body: JSON.stringify({
                action: 'sendFeedback',
                row: Number(row),
                feedbackText: text,
                feedbackHtml: html
            })
        });

        const rawText = await response.text();

        let result;
        try {
            result = JSON.parse(rawText);
        } catch (err) {
            throw new Error('Server returned non-JSON: ' + rawText);
        }

        if (!result.ok) {
            throw new Error(result.error || 'Apps Script send failed.');
        }

        return result;
    }

    return {
        send: sendFeedback
    };
})();

window.QuikDraftSend = QuikDraftSend;