const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        LevelFormat, UnderlineType } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function heading(text, level = 1) {
  const sizes = { 1: 36, 2: 28, 3: 24 };
  const colors = { 1: "1F3864", 2: "2E75B6", 3: "2E75B6" };
  return new Paragraph({
    spacing: { before: level === 1 ? 360 : 200, after: 120 },
    children: [new TextRun({
      text,
      bold: true,
      size: sizes[level],
      color: colors[level],
      font: "Arial"
    })]
  });
}

function qHeading(text) {
  return new Paragraph({
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: "C00000", font: "Arial" })]
  });
}

function subQ(text) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, color: "1F3864", font: "Arial" })]
  });
}

function answerLabel(text) {
  return new Paragraph({
    spacing: { before: 60, after: 40 },
    children: [new TextRun({ text, bold: true, size: 22, color: "375623", font: "Arial" })]
  });
}

function para(text, { bold = false, italic = false, indent = false, size = 22 } = {}) {
  return new Paragraph({
    indent: indent ? { left: 720 } : undefined,
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, bold, italic, size, font: "Arial", color: "000000" })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function numbered(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
    children: []
  });
}

function makeTable(rows, colWidths, headerRow = false) {
  return new Table({
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map((row, ri) => new TableRow({
      children: row.map((cell, ci) => new TableCell({
        borders,
        width: { size: colWidths[ci], type: WidthType.DXA },
        shading: ri === 0 && headerRow
          ? { fill: "2E75B6", type: ShadingType.CLEAR }
          : (ri % 2 === 0 ? { fill: "EBF3FB", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR }),
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({
            text: cell,
            size: 20,
            font: "Arial",
            bold: ri === 0 && headerRow,
            color: ri === 0 && headerRow ? "FFFFFF" : "000000"
          })]
        })]
      }))
    }))
  });
}

function sp(n = 1) {
  return Array(n).fill(null).map(() => new Paragraph({ children: [] }));
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // TITLE
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "ENGLISH I — CLASS IX", bold: true, size: 40, color: "1F3864", font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Board of Secondary Education Karachi (BSEK)", size: 26, color: "2E75B6", font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [new TextRun({ text: "Complete Question Paper — Model Answer Key", bold: true, size: 26, color: "C00000", font: "Arial" })]
      }),

      divider(),

      // ============================================================
      // SECTION B
      // ============================================================
      heading("SECTION B — SHORT ANSWER QUESTIONS", 1),
      new Paragraph({
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "Marks: 40  |  Answer any FIVE of the following in 2–3 sentences each.", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),

      // Q2
      heading("Q.2 — Short Answer Questions (Any 5)", 2),

      // (i)
      subQ("(i) What do you know about the broad groups that existed before the arrival of the Holy Prophet ﷺ in Madinah?"),
      answerLabel("Answer:"),
      para("At the time of the Holy Prophet's (S.A.W) arrival in Madinah, the city had four broad groups of people:"),
      bullet("Muslims — both those originally from Madinah and those who had emigrated from Makkah."),
      bullet("Hypocrites (Munafiqeen) — those who had embraced Islam in name only but were secretly against it."),
      bullet("People of Aws and Khazraj — who were still pagans but were likely to accept Islam."),
      bullet("Jews — who were very large in number and formed an important community in the city."),
      ...sp(1),

      // (ii)
      subQ("(ii) What is the main message of Shah Latif's poetry? Do you agree? Give reason."),
      answerLabel("Answer:"),
      para("The main message of Shah Abdul Latif Bhitai's poetry is love — spiritual, universal, and eternal. His mystic poetry carries a message of love that is acceptable to the entire human race. Yes, I agree with this because his poetry has been translated into Urdu, English, and many other languages, proving that its message is timeless and universal, not limited to Sindh alone."),
      ...sp(1),

      // (iii)
      subQ("(iii) Why did King Hal call the miller \"England's boast\"?"),
      answerLabel("Answer:"),
      para("King Hal called the miller \"England's boast\" because the miller was truly happy and content with his simple life. He worked hard, earned his bread honestly, loved his wife, children, and friends, owed no debt he could not pay, and said: \"I envy nobody, no, not I!\" King Hal himself was sad despite being a king, so he admired the miller's genuine happiness and considered him the pride of England."),
      ...sp(1),

      // (iv)
      subQ("(iv) What are preventive methods for mosquito-borne diseases?"),
      answerLabel("Answer:"),
      para("Since there are no vaccines for these diseases, the following preventive methods should be followed:"),
      bullet("Wear clothes that keep the body fully covered."),
      bullet("Use net coverings while sleeping."),
      bullet("Live in hygienic surroundings."),
      bullet("Ensure there is no stagnant water in the locality — not even in buckets, pots, or cans."),
      bullet("Apply good-quality mosquito repellent lotion or natural repellent products."),
      bullet("Take extra precautions from dawn to dusk."),
      bullet("Seek medical help as soon as symptoms appear."),
      ...sp(1),

      // (v)
      subQ("(v) What role did Professor Anita Ghulam Ali play in the field of education?"),
      answerLabel("Answer:"),
      para("Professor Anita Ghulam Ali (1934–2014) served twice as the Sindh Minister of Education, and also as Minister of Culture, Science and Technology and Youth and Sports. She was the founding manager of the Sindh Education Foundation (SEF), Government of Sindh. Under her leadership, the SEF played a remarkable role in promoting quality education and bringing many out-of-school children — especially in rural Sindh — back to school. The Government of Pakistan recognized her tremendous contributions by awarding her the Pride of Performance and Sitara-e-Imtiaz."),
      ...sp(1),

      // (vi)
      subQ("(vi) What were the three questions that the king wanted to know?"),
      answerLabel("Answer:"),
      para("In the story \"The Secret of Success\" (Unit 7), the king wanted to find answers to three questions:"),
      numbered("What is the most important time to start something?"),
      numbered("What is the most important work to do?"),
      numbered("Who is the most important person?"),
      para("He believed that if he found these answers, he would never fail in life."),
      ...sp(1),

      // (vii)
      subQ("(vii) How did Quaid-e-Azam praise Iqbal's role on his death in a message?"),
      answerLabel("Answer:"),
      para("On the death of Allama Iqbal, Quaid-e-Azam Muhammad Ali Jinnah paid tribute to him in a message, saying: \"To me he was a friend, guide and philosopher, and during the darkest moments through which the Muslim League had to go, he stood like a rock.\" This shows how highly Quaid-e-Azam valued Iqbal's friendship, wisdom, and unwavering support for the cause of Pakistan."),

      divider(),

      // ============================================================
      // Q3
      // ============================================================
      heading("Q.3 — Do As Directed", 2),
      new Paragraph({
        spacing: { before: 0, after: 100 },
        children: [new TextRun({ text: "(Marks: 20)", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),

      // (i) Articles
      subQ("(i) Use Article"),
      answerLabel("Answer:"),
      para("a.  THE Indus is the largest river of Pakistan.", { indent: true }),
      para("    Rule: Names of rivers always take the definite article 'the'.", { indent: true, italic: true }),
      ...sp(1),
      para("b.  Ali is AN intelligent boy in the class.", { indent: true }),
      para("    Rule: 'An' is used before a vowel sound. The word 'intelligent' begins with the vowel sound /ɪ/.", { indent: true, italic: true }),
      ...sp(1),

      // (ii) Prepositions
      subQ("(ii) Use Preposition"),
      answerLabel("Answer:"),
      para("a.  I have been waiting for you FOR two hours.", { indent: true }),
      para("    Rule: 'For' is used with a period/duration of time.", { indent: true, italic: true }),
      ...sp(1),
      para("b.  He is afraid OF spiders.", { indent: true }),
      para("    Rule: The adjective 'afraid' is always followed by the preposition 'of'.", { indent: true, italic: true }),
      ...sp(1),

      // (iii) Voice
      subQ("(iii) Change the Voice"),
      answerLabel("Answer:"),
      makeTable([
        ["Original Sentence", "Changed Voice"],
        ["a. Call the police at once. (Active)", "Let the police be called at once. (Passive)"],
        ["b. Why are you beating him? (Active)", "Why is he being beaten by you? (Passive)"],
        ["c. She is watering the plants. (Active)", "The plants are being watered by her. (Passive)"],
        ["d. English is taught by the teacher. (Passive)", "The teacher teaches English. (Active)"],
      ], [4200, 4800], true),
      ...sp(1),

      // (iv) Narration
      subQ("(iv) Change the Narration"),
      answerLabel("Answer:"),
      makeTable([
        ["Original (Direct Speech)", "Changed (Indirect Speech)"],
        ['a. He said, "I will visit you tomorrow."', "He said that he would visit me/them the next day."],
        ['b. He said, "May you live long!"', "He prayed that I/they might live long."],
        ['c. He said to me, "Hurrah! I have passed the examination."', "He exclaimed with joy and told me that he had passed the examination."],
        ['d. I forbade him to make a noise there.', 'He said to me, "Do not make a noise here."'],
      ], [4200, 4800], true),
      ...sp(1),

      // (v) Formation of Sentences
      subQ("(v) Formation of Sentences"),
      answerLabel("Answer:"),
      makeTable([
        ["#", "Question", "Answer", "Rule/Explanation"],
        ["a", "He (clear/cleared) his all papers in the first go.", "CLEARED", "Past tense required — Simple Past"],
        ["b", 'The word "untidy" is formed using: only suffix / only prefix / suffix and prefix / root word', "ONLY PREFIX", "un- is a prefix; 'tidy' is the root. No suffix is added."],
        ["c", 'Correct suffix to form a noun from "agree": ment / nace / all / ive', "-MENT → agreement", "The noun form of 'agree' is 'agreement'."],
        ["d", "The train had left station on time. (Change into Future Continuous Tense)", "The train will be leaving the station on time.", "Future Continuous = will be + V-ing"],
        ["e", "Ali will be reciting the Holy Quran. (Change into Past Continuous)", "Ali was reciting the Holy Quran.", "Past Continuous = was/were + V-ing"],
        ["f", "Mother cooks meal for her children. (Negative)", "Mother does not cook meal for her children.", "Simple Present Negative = do/does + not + V1"],
        ["g", "sana humera and sadia please come here (Punctuate)", "Sana, Humera and Sadia, please come here.", "Capitalize names; comma after each name and after 'Sadia'."],
        ["h", "He is superior than her. (Correct)", "He is superior to her.", "'Superior' takes 'to', not 'than'."],
      ], [360, 2800, 2500, 2840], true),

      divider(),

      // ============================================================
      // Q4
      // ============================================================
      heading("Q.4 — Indicate the Part of Speech of the Underlined Words", 2),
      new Paragraph({
        spacing: { before: 0, after: 100 },
        children: [new TextRun({ text: "(Marks: 5)", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),
      answerLabel("Answer:"),
      makeTable([
        ["#", "Sentence (Underlined Word)", "Word", "Part of Speech", "Reason"],
        ["(i)", "The cat is UNDER the table.", "under", "Preposition", "Shows the relationship/position of the cat to the table."],
        ["(ii)", "WOW! That is amazing.", "Wow", "Interjection", "Expresses a sudden emotion or exclamation."],
        ["(iii)", "She sings song very SWEETLY.", "sweetly", "Adverb", "Modifies the verb 'sings'; tells how she sings."],
        ["(iv)", "He ran FAST toward the market.", "fast", "Adverb", "Modifies the verb 'ran'; tells how he ran."],
        ["(v)", "Salma BELONGED to Hyderabad.", "belonged", "Verb", "Expresses action/state in the past."],
      ], [360, 3000, 1000, 1500, 2640], true),

      divider(),

      // ============================================================
      // Q5
      // ============================================================
      heading("Q.5 — Translate the Paragraph into Urdu / Sindhi", 2),
      new Paragraph({
        spacing: { before: 0, after: 100 },
        children: [new TextRun({ text: "(Marks: 3)", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),
      qHeading("Original Paragraph (English):"),
      new Paragraph({
        spacing: { before: 60, after: 120 },
        shading: { fill: "FFF9E6", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "F0A500", space: 10 } },
        indent: { left: 360 },
        children: [new TextRun({ text: "A young boy lived in a small village near a river. Every morning, he helped his father in the fields before going to school. One day, heavy rain caused the river to overflow and flood the village. The boy quickly gathered his books and helped his neighbours move to safer places. His bravery and quick thinking saved many lives.", size: 22, font: "Arial", italic: true })]
      }),
      qHeading("Urdu Translation:"),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "ایک نوجوان لڑکا ایک دریا کے قریب ایک چھوٹے سے گاؤں میں رہتا تھا۔ ہر صبح، وہ اسکول جانے سے پہلے کھیتوں میں اپنے باپ کی مدد کرتا تھا۔ ایک دن، بھاری بارش کی وجہ سے دریا بہہ نکلا اور گاؤں میں سیلاب آ گیا۔ لڑکے نے فوری طور پر اپنی کتابیں اکٹھی کیں اور اپنے پڑوسیوں کو محفوظ جگہوں پر جانے میں مدد کی۔ اس کی بہادری اور فوری سوچ نے بہت سی جانیں بچائیں۔", size: 24, font: "Arial" })]
      }),

      divider(),

      // ============================================================
      // SECTION C
      // ============================================================
      heading("SECTION C — DESCRIPTIVE ANSWER QUESTIONS", 1),
      new Paragraph({
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "Marks: 40", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),

      // Q6
      heading("Q.6 — Fill in the Blanks (5 Marks)", 2),
      qHeading("Question:"),
      para("Fill in the blanks according to the contextual reference from the options provided in the box:", { italic: true }),
      new Paragraph({
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: "Box: renowned | tireless | nationwide | philanthropist | recognition", bold: true, size: 22, font: "Arial", color: "1F3864" })]
      }),
      para("She is a ________ who is also trained as a professional nurse. She is the widow of Abdul Sattar Edhi, she was a ________ social worker. She now heads the Edhi Foundation which provides ________ support and care to the public, such as shelter, homes and ambulance services. She has received the Hilal-e-Imtiaz from the Government of Pakistan in ________ of her ________ efforts for the betterment of the society.", { italic: true }),
      answerLabel("Answer (Completed Paragraph):"),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        shading: { fill: "E8F5E9", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "27AE60", space: 10 } },
        indent: { left: 360 },
        children: [
          new TextRun({ text: "She is a ", size: 22, font: "Arial" }),
          new TextRun({ text: "philanthropist", bold: true, size: 22, font: "Arial", color: "C00000" }),
          new TextRun({ text: " who is also trained as a professional nurse. She is the widow of Abdul Sattar Edhi, she was a ", size: 22, font: "Arial" }),
          new TextRun({ text: "tireless", bold: true, size: 22, font: "Arial", color: "C00000" }),
          new TextRun({ text: " social worker. She now heads the Edhi Foundation which provides ", size: 22, font: "Arial" }),
          new TextRun({ text: "nationwide", bold: true, size: 22, font: "Arial", color: "C00000" }),
          new TextRun({ text: " support and care to the public, such as shelter, homes and ambulance services. She has received the Hilal-e-Imtiaz from the Government of Pakistan in ", size: 22, font: "Arial" }),
          new TextRun({ text: "recognition", bold: true, size: 22, font: "Arial", color: "C00000" }),
          new TextRun({ text: " of her ", size: 22, font: "Arial" }),
          new TextRun({ text: "renowned", bold: true, size: 22, font: "Arial", color: "C00000" }),
          new TextRun({ text: " efforts for the betterment of the society. She has also been named the Mother of Pakistan.", size: 22, font: "Arial" }),
        ]
      }),

      divider(),

      // ============================================================
      // Q7 — ALL ESSAYS
      // ============================================================
      heading("Q.7 — Essay Writing (10 Marks)", 2),
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "Write an essay of 120–150 words on ANY ONE of the following topics:", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),
      new Paragraph({
        spacing: { before: 40, after: 200 },
        children: [new TextRun({ text: "All five options are answered below for complete practice.", bold: true, size: 22, color: "C00000", font: "Arial" })]
      }),

      // Essay 1
      heading("Essay Option (i): The Importance of Computers in Our Lives", 3),
      new Paragraph({
        spacing: { before: 80, after: 300 },
        shading: { fill: "F0F4FF", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "The Importance of Computers in Our Lives\n\nIn the modern world, computers have become an essential part of daily life. From education and communication to business and entertainment, computers are used in almost every field. They have made life faster, easier, and more organized.\n\nIn the field of education, computers give students access to vast amounts of information through the internet. Students can attend online classes, research topics, and submit assignments digitally. In hospitals, computers help doctors maintain patient records and conduct research. In business, computers are used for accounting, communication, and management.\n\nComputers have also revolutionized communication. Through emails, video calls, and social media, people can connect with others anywhere in the world within seconds. They have also created new job opportunities in fields like programming, graphic design, and data science.\n\nHowever, excessive use of computers can lead to health problems such as eye strain and poor posture. Therefore, we must use computers wisely and in balance. In conclusion, computers are an indispensable tool of the modern age, and learning to use them is a necessity, not a luxury.", size: 22, font: "Arial" })]
      }),

      // Essay 2
      heading("Essay Option (ii): Ideal Personality", 3),
      new Paragraph({
        spacing: { before: 80, after: 300 },
        shading: { fill: "F0F4FF", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "An Ideal Personality\n\nAn ideal personality is one that combines good character, strong moral values, and positive behaviour. A person with an ideal personality is honest, humble, hardworking, and kind to others. They treat every person with respect, regardless of religion, status, or background.\n\nSuch a person is also disciplined and punctual. They fulfil their responsibilities towards their family, their society, and their country. When faced with difficulties, they do not give up, but instead face challenges with courage, patience, and wisdom.\n\nAn ideal person is never arrogant. Despite their achievements, they remain humble and always give credit to others. They listen carefully before speaking and always choose the truth over convenience.\n\nIn my view, our Holy Prophet Muhammad (S.A.W) is the greatest example of an ideal personality in all of human history. His life teaches us patience, justice, compassion, and love for all of humanity. If we sincerely follow his example, we too can develop ideal personalities and make a positive difference in the world.", size: 22, font: "Arial" })]
      }),

      // Essay 3
      heading("Essay Option (iii): A Visit to a Historical Place", 3),
      new Paragraph({
        spacing: { before: 80, after: 300 },
        shading: { fill: "F0F4FF", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "A Visit to a Historical Place\n\nLast winter, our school arranged a visit to the ancient city of Thatta, located in Sindh, Pakistan. It was one of the most memorable experiences of my life. Thatta was once the capital of the Samma dynasty and later flourished under Mughal rule, making it a city of great historical importance.\n\nWe first visited Makli Hill, one of the largest funerary sites in the world. The tombs and mausoleums there date from the 14th to the 18th centuries and reflect a fascinating blend of Islamic, Persian, and Hindu architectural styles. The grandeur of the monuments left us speechless.\n\nNext, we visited the Shah Jahan Mosque, which is famous for its stunning blue-tiled domes and beautiful calligraphy. The intricate design and craftsmanship were truly remarkable.\n\nThe visit taught me the importance of preserving our heritage. Thatta is now a UNESCO World Heritage Site, but it faces threats from neglect and environmental damage. As young Pakistanis, it is our duty to protect and promote our rich cultural history. I returned home with a deep sense of pride and responsibility.", size: 22, font: "Arial" })]
      }),

      // Essay 4
      heading("Essay Option (iv): Our Past is Better than Present", 3),
      new Paragraph({
        spacing: { before: 80, after: 300 },
        shading: { fill: "F0F4FF", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "Our Past is Better than Present\n\nMany people believe that the past was a better time to live in, and there are several reasons to support this view. In the past, life was simpler, slower, and more peaceful. People had strong relationships with their families and communities. Neighbours knew and helped each other, and there was a real sense of unity and belonging.\n\nValues such as honesty, respect, and hard work were central to everyday life. Children were raised with discipline and a clear moral compass. There was less crime, less pollution, and less stress. People enjoyed pure natural food, clean air, and a slower pace of life that allowed for rest and reflection.\n\nHowever, it is also fair to say that the present has its own advantages. Modern medicine has greatly increased life expectancy. Technology has made communication, education, and travel far easier and faster.\n\nIn conclusion, while the present offers convenience and progress, the past offered stronger human connections and simpler values. Perhaps the best path is to take the wisdom and values of the past and apply them to the opportunities of the present. A society that respects its roots while embracing progress is the most likely to succeed.", size: 22, font: "Arial" })]
      }),

      // Email Option
      heading("Essay Option — OR: Write an E-mail to a Friend about a Funny Incident", 3),
      new Paragraph({
        spacing: { before: 80, after: 300 },
        shading: { fill: "FFF9E6", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "F0A500", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "To: Ahmed@gmail.com\nFrom: YourName@gmail.com\nSubject: You Won't Believe What Happened Today!\nDate: 22nd June, 2026\n\nDear Ahmed,\n\nI hope this email finds you in good health and great spirits! I am writing to tell you about the most hilarious thing that happened to me today — I simply cannot stop laughing every time I think about it!\n\nThis morning, I woke up late and rushed to get ready for school. In my hurry, I accidentally put on two different shoes — one black and one brown! I did not realize this until I arrived at school and my best friend Bilal burst out laughing. I had walked all the way from home, taken a rickshaw, and even greeted the school guard — all with mismatched shoes! The whole class could not stop laughing, and even our class teacher joined in.\n\nThe most embarrassing part was when our principal walked past and noticed as well! Thankfully, she just smiled and said, \"At least they were the same size!\"\n\nI had to borrow a marker from the art room to make my brown shoe look black for the rest of the day!\n\nI hope this little story made you smile too. Write back soon!\n\nYours truly,\n[Your Name]", size: 22, font: "Arial" })]
      }),

      divider(),

      // ============================================================
      // Q8
      // ============================================================
      heading("Q.8 — Application Writing (10 Marks)", 2),
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "Write an application to your Headmaster/Headmistress requesting a half day leave to take your mother to the hospital.", size: 22, italic: true, color: "555555", font: "Arial" })]
      }),
      answerLabel("Answer — Application:"),
      new Paragraph({
        spacing: { before: 80, after: 60 },
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "Date: 22nd June, 2026", size: 22, font: "Arial" })]
      }),
      para("The Headmaster / Headmistress,"),
      para("[Name of School],"),
      para("Karachi."),
      ...sp(1),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Subject: Application for Half Day Leave", bold: true, size: 22, font: "Arial", underline: { type: UnderlineType.SINGLE } })]
      }),
      ...sp(1),
      para("Respected Sir / Madam,"),
      ...sp(1),
      para("With due respect, I beg to state that I am a regular student of Class IX at your esteemed school. I am writing this application to request a half day leave for today."),
      ...sp(1),
      para("My mother is not feeling well and requires an immediate medical consultation. Unfortunately, there is no other family member available to accompany her to the hospital at this time. Therefore, it is necessary for me to go with her personally to ensure she receives proper care and attention."),
      ...sp(1),
      para("I kindly request you to grant me half day leave from school today so that I may take my mother to the doctor. I assure you that I will make up for all missed classwork and notes as soon as I return."),
      ...sp(1),
      para("I shall be very thankful to you for this kind consideration."),
      ...sp(2),
      para("Yours obediently,"),
      para("[Your Full Name]"),
      para("Class IX, Roll No. [___]"),
      para("Contact: [Phone Number]"),

      ...sp(1),
      new Paragraph({
        spacing: { before: 120, after: 100 },
        children: [new TextRun({ text: "— OR — Write a Diary Entry about a Day at the Seaside", bold: true, size: 24, color: "C00000", font: "Arial" })]
      }),
      answerLabel("Answer — Diary Entry:"),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        shading: { fill: "FFF9E6", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "F0A500", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "Monday, 22nd June, 2026\n\nDear Diary,\n\nToday was absolutely wonderful — one of the best days I have had in a very long time! Our school organized a trip to the seaside, and I could not have been more excited. We left early in the morning and arrived at Clifton Beach just as the sun was rising over the sea. The sight of the golden waves glittering in the morning light was breathtaking.\n\nMy classmates and I played in the shallow water, collected seashells, and built sandcastles. Some of the boys challenged each other to races along the shore, while others sat quietly and watched the waves. Our teacher, Sir Kamran, organized a small quiz competition on the beach and gave small prizes to the winners. I won a bookmark — which I am going to keep as a memory of today!\n\nWe had lunch together on the sand — sandwiches, juice, and biscuits. Everything tasted especially delicious in the sea breeze! We also watched the boats sail in the distance and talked about far-off places we would love to visit one day.\n\nAs we headed back in the evening, everyone was tired but happy. The sound of the waves, the smell of the salty air, and the laughter of my friends — I will carry these memories with me for a long time.\n\nGoodnight, Diary. Today reminded me how beautiful the world is.\n\nYours,\n[Your Name]", size: 22, font: "Arial" })]
      }),

      divider(),

      // ============================================================
      // Q9
      // ============================================================
      heading("Q.9 — Reading Comprehension: Thatta (15 Marks)", 2),
      qHeading("Passage:"),
      new Paragraph({
        spacing: { before: 60, after: 160 },
        shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "555555", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "The ancient city of Thatta, located in Sindh, was once a vibrant centre of trade and culture. For centuries, it served as the capital of the Samma dynasty and later as a prominent city under the Mughals. Its grand necropolis, Makli Hill, is one of the largest funerary sites in the world, with tombs and monuments dating from the 14th to the 18th centuries. The architecture reflects a fusion of Islamic, Persian, and Hindu influences, showcasing the skill of local artisans. Today, Thatta is a UNESCO World Heritage Site, but it faces challenges such as neglect, encroachment, and environmental decay.\n\nEfforts are underway to preserve this heritage. The government, in collaboration with international organizations, has initiated restoration projects. Local communities are being involved in protecting the site, and awareness campaigns aim to highlight its importance. However, more sustained funding and stricter enforcement of conservation laws are needed. The passage stresses that losing Thatta would mean losing a priceless chapter of Pakistan's history.", size: 21, font: "Arial", italic: true, color: "333333" })]
      }),

      // (i)
      subQ("(i) Which dynasty made Thatta its capital, and what is the name of the famous funerary site there? (2 marks)"),
      answerLabel("Answer:"),
      para("The Samma dynasty made Thatta its capital. The famous funerary site located there is called Makli Hill, which is one of the largest funerary sites in the entire world, with tombs and monuments dating from the 14th to the 18th centuries."),
      ...sp(1),

      // (ii)
      subQ("(ii) What does the architecture reflect? What challenges are being faced by UNESCO in this concern? (2 marks)"),
      answerLabel("Answer:"),
      para("The architecture of Thatta reflects a beautiful fusion of Islamic, Persian, and Hindu influences, showcasing the extraordinary skill of local artisans. As for the challenges, Thatta faces serious problems including neglect, encroachment (illegal occupation of land), and environmental decay. These threats put the site at risk of being permanently damaged or lost."),
      ...sp(1),

      // (iii)
      subQ("(iii) Make an adjective of: Culture (1 mark)"),
      answerLabel("Answer:"),
      para("Adjective of 'Culture':  CULTURAL"),
      para("(e.g., \"Thatta is a city of great cultural importance.\")", { italic: true }),
      ...sp(1),

      // (iv)
      subQ("(iv) Find a word from the passage that means: Hundreds of years / With the cooperation / Graveyard / To keep safe / Compulsion (5 marks)"),
      answerLabel("Answer:"),
      makeTable([
        ["Meaning Given", "Word Found in Passage"],
        ["Hundreds of years", "Centuries"],
        ["With the cooperation", "In collaboration"],
        ["Graveyard", "Necropolis"],
        ["To keep safe", "Preserve"],
        ["Compulsion", "Enforcement"],
      ], [4200, 4800], true),
      ...sp(1),

      // (v)
      subQ("(v) Make a summary of the passage with a suitable title. (5 marks)"),
      answerLabel("Answer:"),
      new Paragraph({
        spacing: { before: 60, after: 80 },
        children: [new TextRun({ text: "Title: Thatta — A Priceless Heritage Under Threat", bold: true, size: 24, font: "Arial", color: "1F3864" })]
      }),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        shading: { fill: "EBF3FB", type: ShadingType.CLEAR },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 10 } },
        indent: { left: 360, right: 360 },
        children: [new TextRun({ text: "Thatta is an ancient city in Sindh, Pakistan, which once served as the capital of the Samma dynasty and flourished under Mughal rule. It is home to Makli Hill, one of the world's largest funerary sites, featuring magnificent tombs that blend Islamic, Persian, and Hindu architectural styles. Today, Thatta holds the status of a UNESCO World Heritage Site. However, it faces serious threats from neglect, illegal encroachment, and environmental damage. The government, working together with international organizations, has begun restoration efforts and is involving local communities in the protection of the site. Despite this progress, more consistent funding and the strict enforcement of conservation laws are urgently needed. Losing Thatta would mean losing an irreplaceable part of Pakistan's history and cultural identity.", size: 22, font: "Arial" })]
      }),

      divider(),

      // Footer note
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 0 },
        children: [new TextRun({ text: "— END OF ANSWER KEY —   |   Prepared from BSEK Class IX English Book", size: 18, italic: true, color: "888888", font: "Arial" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/English_IX_Complete_Paper_Solution.docx", buffer);
  console.log("Done!");
});