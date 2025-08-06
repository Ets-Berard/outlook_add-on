import OPENAI_API_KEY from './config.js';

Office.onReady(() => {});

async function processEmail(action) {
    Office.context.mailbox.item.body.getAsync("text", async (result) => {
        const tone = document.getElementById("tone").value;
        const basePrompt = {
            summary: `Résume cet e-mail de manière ${tone} :\n\n`,
            reply: `Rédige une réponse ${tone} à cet e-mail :\n\n`,
            rephrase: `Reformule ce message de manière ${tone} :\n\n`
        };

        const prompt = basePrompt[action] + result.value;

        const response = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${OPENAI_API_KEY}`
            },
            body: JSON.stringify({
                model: "gpt-4o-mini",
                messages: [{ role: "user", content: prompt }],
                max_tokens: 500
            })
        });

        const data = await response.json();
        document.getElementById("output").value = data.choices[0].message.content;
    });
}

window.insertDraft = function insertDraft() {
    const text = document.getElementById("output").value;
    if (!text) {
        alert("Aucun texte généré.");
        return;
    }

    Office.context.mailbox.item.body.setAsync(text, { coercionType: "html" }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            alert("Texte inséré dans le brouillon !");
        } else {
            alert("Erreur lors de l’insertion.");
        }
    });
}

window.processEmail = processEmail;
