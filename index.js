


let pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";

pptx.defineSlideMaster({
  title: "Default_ZMC",
  background: { color: "FFFFFF" },
  objects: [
    {
      image: { x: 0.5, y: 0.26, w: 0.39, h: 0.77, path: "https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/zmc_logo.png" },
    },
    {
      line: { x: 0.5, y: 5.34, w: 9, h: 0, line: "808EAA" },
    },
    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.5, y: 0.22, w: 9, h: 0.94 },
        text: "Click to edit the title text format",
      },
    },
  ],
});

function preloadImages(urls, callback) {
  let loadedCounter = 0;
  let images = [];

  for (let i = 0; i < urls.length; i++) {
      images[i] = new Image();
      images[i].onload = function() {
          loadedCounter++;
          if (loadedCounter === urls.length) {
              callback(); // Call the callback function when all images are loaded
          }
      };
      images[i].src = urls[i];
  }
}


let imageUrls = [
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/logos/color/image%2079.jpg',
  // Add more URLs as needed
];

document.getElementById('loadButton').addEventListener('click', function() {
  preloadImages(imageUrls, function() {
      console.log("All images preloaded");
      createLogoSlide(imageUrls);
  });
});


function createLogoSlide(imageUrls) {
  console.log(imageUrls)
  
  let slide = pptx.addSlide({ masterName: "Default_ZMC" });
  let gridX = 0.5, gridY = 1.63, gridW = 9, gridH = 4;
  let cols = 10, rows = 5;
  let horizontalMargin = 0.05; // Horizontal margin on each side
  let verticalMargin = 0.10;   // Vertical margin on each side
  let cellWidth = gridW / cols;
  let cellHeight = gridH / rows;

  // Effective image size after accounting for margins
  let imageWidth = cellWidth - 2 * horizontalMargin;
  let imageHeight = cellHeight - 2 * verticalMargin;

  // Add images in a 10x5 grid
  for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
          let index = i * cols + j;
          if (index < imageUrls.length) {
              slide.addImage({
                  path: imageUrls[index],
                  x: gridX + j * cellWidth + horizontalMargin,
                  y: gridY + i * cellHeight + verticalMargin,
                  w: imageWidth,
                  h: imageHeight
              });
          }
      }
  }

  // Write the file
  pptx.writeFile({ fileName: 'ImageGridPresentation.pptx' });
}


let projectExamples = [
  {heading: 'Project 1', rt: 'rich text', company_name_neutral: 'kjkdsf jkhsdf hs'},
  {heading: 'Project 2', rt: 'rich text', company_name_neutral: 'kjkdsf jkhsdf hs'},
  {heading: 'Project 3', rt: 'rich text', company_name_neutral: 'kjkdsf jkhsdf hs'},
  {heading: 'Project 4', rt: 'rich text', company_name_neutral: 'kjkdsf jkhsdf hs'},
  {heading: 'Project 5', rt: 'rich text', company_name_neutral: 'kjkdsf jkhsdf hs'},
  {heading: 'Project 6', rt: 'rich text', company_name_neutral: 'kjkdsf jkhsdf hs'},
];

document.getElementById('exampleButton').addEventListener('click', function() {
  createProjectExamples(projectExamples);

  function createProjectExamples(projectExamples) {
    let slide = pptx.addSlide({ masterName: "Default_ZMC" });
    let gridX = 0.5, gridY = 1.63, gridW = 9, gridH = 4;
    let cols = 5, rows = 4;
    let horizontalMargin = 0.05; // Horizontal margin on each side
    let verticalMargin = 0.05;   // Vertical margin on each side
    let cellWidth = gridW / cols;
    let cellHeight = gridH / rows;
    let cellHeightSmall = 0.23;

    let elementWidth = cellWidth - 2 * horizontalMargin;
    let elementHeight = cellHeight - 2 * verticalMargin;

    for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
            let index = i * cols + j;
            if (index < projectExamples.length) {
                let yPosition = gridY + i * cellHeight + verticalMargin;
                
                // Add rich text (full cell height)
                slide.addText(projectExamples[index].rt, {
                  shape: pptx.ShapeType.ROUNDED_RECTANGLE,
                  x: gridX + j * cellWidth + horizontalMargin,
                  y: yPosition,
                  w: elementWidth,
                  h: elementHeight,
                  rectRadius: 10, 
                  fill: { color: "e4e6ec" },
                  fontSize: 7,
                    color: "000000",
                });
                // Add heading
                slide.addText(projectExamples[index].heading, {
                  shape: pptx.ShapeType.roundRect,
                  x: gridX + j * cellWidth + horizontalMargin,
                  y: yPosition, // Align to the top of the cell
                  w: elementWidth,
                  h: cellHeightSmall, 
                  rectRadius: 10, 
                  fill: { color: "535d76" },
                  fontSize: 7,
                  color: "FFFFFF",
                  fontWeight: 'bold',
                });
                // Add company name
                slide.addText(projectExamples[index].company_name_neutral, {
                  shape: pptx.ShapeType.roundRect,
                  x: gridX + j * cellWidth + horizontalMargin,
                  y: yPosition + elementHeight - cellHeightSmall, // Align to the bottom of the cell
                  w: elementWidth,
                  h: cellHeightSmall, 
                  rectRadius: 10, 
                  fill: { color: "ffffff" },
                  fontSize: 7,
                  fontWeight: 'bold',
                  color: "535d76",
                });
            }
        }
    }

    pptx.writeFile({ fileName: 'projectExample.pptx' });
  }});





let referenceProject = {company_name_neutral: 'kjkdsf jkhsdf hs', project: 'Project 1', goal: 'rich text', solution: 'rich text'};


document.getElementById('referenceButton').addEventListener('click', function() {
  createReferenceProject(referenceProject);
  let pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";

pptx.defineSlideMaster({
  title: "REFERENCE_ZMC",
  background: { color: "FFFFFF" },
  objects: [
    {
      image: { x: 0.5, y: 0.26, w: 0.39, h: 0.77, path: "https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/zmc_logo.png" },
    },
    {
      line: { x: 0.5, y: 5.34, w: 9, h: 0, line: "808EAA" },
    },
    {
      rect: {
        x: 0.5,
        y: 1.23,
        w: 1.92,
        h: 2.25,
        fill: { color: "e4e6ec" },
      }
    },
    {
      roundRect: {
        x: 0.54,
        y: 1.26,
        w: 0.3,
        h: 0.3,
        rectRadius: 200, 
        fill: { color: "b3001e" },
      }
    },
    {
      image: { 
        path: "https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/briefcase.png",
        x: 0.59,
        y: 1.3,
        w: 0.2,
        h: 0.2,
      }},
    {
      text: { text: "Unternehmen", options: {
      x: 0.83,
      y: 1.24,
      w: 1.43,
      h: 0.32,  
      fontSize: 10,
      fontWeight: 'bold',
      color: "000000",
      }},
    },
    {
      rect: {
        x: 2.5,
        y: 1.23,
        w: 3.17,
        h: 2.25,
        rectRadius: 2, 
        fill: { color: "e4e6ec" },
      }
    },
    {
      roundRect: {
        x: 2.54,
        y: 1.26,
        w: 0.3,
        h: 0.3,
        rectRadius: 200, 
        fill: { color: "b3001e" },
      }
    },
    {
      image: { 
        path: "https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/target.png",
        x: 2.58,
        y: 1.31,
        w: 0.2,
        h: 0.2,
      }},
    {
      text: {text: "Projekt", options:{
      x: 2.83,
      y: 1.24,
      w: 1.74,
      h: 0.32,  
      fontSize: 10,
      fontWeight: 'bold',
      color: "000000",
    }},
    },
    {
      rect: {
        x: 5.75,
        y: 1.23,
        w: 3.75,
        h: 2.25,
        rectRadius: 2, 
        fill: { color: "e4e6ec" },
      }
    },
    {
      roundRect: {
        x: 5.78,
        y: 1.26,
        w: 0.3,
        h: 0.3,
        rectRadius: 200, 
        fill: { color: "b3001e" },
      }
    },
    {
      image: { 
        path: "https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/process.png",
        x: 5.82,
        y: 1.28,
        w: 0.25,
        h: 0.25,
      }},
    {
      text: {text: "Ziel", options:{
        x: 6.1,
        y: 1.24,
        w: 1.44,
        h: 0.32,  
        fontSize: 10,
        fontWeight: 'bold',
        color: "000000",
      }},
    },
    {
      rect: {
        x: 0.5,
        y: 3.56,
        w: 9,
        h: 1.58,
        rectRadius: 2,
        line: { color: "e4e6ec" },
        fill: { color: "e4e6ec" },
      }
    },
    {
      roundRect: {
        x: 0.54,
        y: 3.61,
        w: 0.3,
        h: 0.3,
        rectRadius: 200, 
        fill: { color: "b3001e" },
      }
    },
    {
      image: {
        path: "https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/solution.png",
        x: 0.58,
        y: 3.63,
        w: 0.21,
        h: 0.25,
      }
    },
    {
      text: {text: "Lösung", options:{
        x: 0.83,
        y: 3.57,
        w: 1.43,
        h: 0.32,  
        fontSize: 10,
        fontWeight: 'bold',
        color: "000000",
      }},
    },

    {
      placeholder: {
        options: { name: "title", type: "title", x: 0.5, y: 0.22, w: 9, h: 0.94 },
        text: "Click to edit the title text format",
      },
    },
  ],
});


  function createReferenceProject(referenceProject) {
    let slide = pptx.addSlide({ masterName: "REFERENCE_ZMC" });
    
    // card 1
    slide.addText(referenceProject.company_name_neutral, {
      x: 0.54,
      y: 1.54,
      w: 1,
      h: 0.3,  
      fontSize: 9,
      color: "000000",
    });

    // card 2
    slide.addText(referenceProject.project, {
      x: 2.54,
      y: 1.54,
      w: 3,
      h: 2,  
      fontSize: 9,
      color: "000000",
    });

    // card 3
    slide.addText(referenceProject.goal, {
      x: 5.8,
      y: 1.54,
      w: 3.5,
      h: 2,  
      fontSize: 9,
      color: "000000",
    });

    // card 4
    slide.addText(referenceProject.solution, {
      x: 0.54,
      y: 3.87,
      w: 8.5,
      h: 1.2,
      fontSize: 9,
      color: "000000",
    });

    pptx.writeFile({ fileName: 'referenceProject.pptx' });
  }
});


const statement = [
  {rt: 'rich text', company_name: 'Test GmbH', name: 'Test Muster', position: 'CFO', image: 'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/zmc_logo.png'},
  {rt: 'rich text', company_name: 'Test Firma', name: 'Muster Test', position: 'CEO', image: 'https://tszpwpclmrfglzjncrjb.supabase.co/storage/v1/object/public/assets/zmc_logo.png'},
];

document.getElementById('statementButton').addEventListener('click', function() { 
  createStatments(statement);

  function createStatments(statement) {
  let slide = pptx.addSlide({ masterName: "Default_ZMC" });
            
  // First statement
  slide.addText("«", {
    x: 2.52,
    y: 1,
    w: 0.39,
    h: 0.5,
    fontSize: 24,
    color: "b3001e",
  });
  slide.addText("»", {
    x: 9.32,
    y: 2.54,
    w: 0.39,
    h: 0.5,
    fontSize: 24,
    color: "b3001e",
  });

  slide.addText(statement[0].rt, {
    x: 2.74,
    y: 1.06,
    w: 7,
    h: 2.21,
    fontSize: 10.5,
    color: "000000",
  });

  slide.addText(statement[0].name + '\n' + statement[0].position + ' ' + statement[0].company_name, {
    x: 0.46,
    y: 2.55,
    w: 2.34,
    h: 0.45,
    fontSize: 10.5,
    color: "000000",
    valign: 'top',
    align: 'center',
  });

  slide.addImage({ path: statement[0].image, x:1, y: 1.3, w: 1.17, h: 1.17, rounding: true });

  // Seconde statement
  slide.addText("«", {
    x: 0.41,
    y: 3.17,
    w: 0.38,
    h: 0.5,
    fontSize: 24,
    color: "b3001e",
  });
  slide.addText("»", {
    x: 7.24,
    y: 4.7,
    w: 0.39,
    h: 0.5,
    fontSize: 24,
    color: "b3001e",
  });

  slide.addText(statement[1].rt, {
    x: 0.58,
    y: 3.45,
    w: 7,
    h: 1.69,
    fontSize: 10.5,
    color: "000000",
  });

  slide.addText(statement[1].name + '\n' + statement[1].position + ' ' + statement[1].company_name, {
    x: 7.31,
    y: 4.53,
    w: 2.3,
    h: 0.6,
    fontSize: 10.5,
    color: "000000",
    valign: 'top',
    align: 'center',
  });

  slide.addImage({ path: statement[1].image, x:7.91, y: 3.31, w: 1.17, h: 1.17, rounding: true });
  pptx.writeFile({ fileName: 'statments.pptx' });
}});



function parseMarkdown(markdownText) {
  // Split markdownText into lines
  const lines = markdownText.split('\n');
  const pptxTextObjects = [];

  // Regular expressions for markdown patterns
  const boldRegex = /\*\*(.*?)\*\*/g;
  const italicRegex = /\*(.*?)\*/g;
  const bulletListRegex = /^- (.*)/;
  const orderedListRegex = /^\d+\. (.*)/;

  lines.forEach(line => {
      let textOptions = { text: "", options: {} };
      let match;

      // Check for bullet list
      if (match = line.match(bulletListRegex)) {
          textOptions.options.bullet = { type: 'bullet' };
          line = match[1];
      }
      // Check for ordered list
      else if (match = line.match(orderedListRegex)) {
          textOptions.options.bullet = { type: 'number' };
          line = match[1];
      }

      // Nested formatting for bold and italic
      let formattedText = [];
      let lastIndex = 0;

      // Bold formatting
      while ((match = boldRegex.exec(line)) !== null) {
          if (match.index > lastIndex) {
              formattedText.push({ text: line.substring(lastIndex, match.index) });
          }
          formattedText.push({ text: match[1], options: { bold: true } });
          lastIndex = match.index + match[0].length;
      }

      // Italic formatting
      if (formattedText.length > 0) {
          formattedText = formattedText.flatMap(fragment => {
              if (fragment.options && fragment.options.bold) {
                  return fragment;
              }
              return fragment.text.split(italicRegex).map((part, index) => 
                  index % 2 === 1 ? { text: part, options: { italic: true } } : { text: part }
              );
          });
      } else {
          line.substring(lastIndex).split(italicRegex).forEach((part, index) => {
              if (index % 2 === 1) {
                  formattedText.push({ text: part, options: { italic: true } });
              } else if (part) {
                  formattedText.push({ text: part });
              }
          });
      }

      if (formattedText.length === 0) {
          formattedText.push({ text: line.substring(lastIndex) });
      }

      pptxTextObjects.push(...formattedText.map(fragment => ({
          text: fragment.text,
          options: { ...textOptions.options, ...fragment.options }
      })));
  });

  return pptxTextObjects;
}


const markdownText = `
khwebhfkwef

  - kjwenfjk nwef
  - wekfnkwe
  - wemnfkwen
  - wef
  
  **kjwen fjkenwf**`;



let pptxTextObjects = parseMarkdown(markdownText);
console.log(pptxTextObjects);
