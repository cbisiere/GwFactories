/* test functions to use from the Apps Script editor */

function test() {
  DocApp.run();
}


function t() {
  const folder = LF.DriveHelper.getFirstParent(SpreadsheetApp.getActiveSpreadsheet());
  createSampleDocument(folder);
}


function createSampleModel(lang) {
  const doc = DocumentApp.create('Modèle (démonstration)');
  const body = doc.getBody();

  const header = body.appendParagraph("Modèle de démonstration");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('Ceci est un modèle de démonstration du module complémentaire “Document factory”.');
  body.appendParagraph('Le table de données contient un champ “Champ 1”.');

  const cells = [
    [
      'Fonctionnalité',
      'Exemple',
      'Résultat',
    ],
    [
      'Remplacement d’un champ par une valeur.',
      'Valeur : <<Champ 1>>',
      'Remplacement du motif par la valeur du champ. Si le champ n’existe pas dans le fichier de données, le motif est laissé tel quel.',
    ],
    [
      'Suppression du motif quand le champ n’existe pas dans le fichier de données.',
      'Valeur : <<Champ 1?>>',
      'Si le champ n’existe pas dans le fichier de données, le motif est supprimé.',
    ],
    [
      'Conservation de la mise en forme lors du remplacement.',
      'Valeur : <<Champ 1>>',
      'Le texte de remplacement sera en gras.',
    ],
    [
      'Suppression d’un paragraphe vide après remplacement.',
      '<<Champ 1>>\nParagraphe suivant.',
      'Si après substitution du motif le premier paragraphe est vide, ce paragraphe vide est supprimé et “Paragraphe suivant.” devient le premier paragraphe de ce texte.',
    ],
    [
      'Usage de texte supplémentaire autour des chevrons internes et suppression du motif quand le champ est vide.',
      '<Valeur : <Champ 1>... >',
      'Si le champ est vide, le motif (tout ce qui est compris entre les chevrons externes) est supprimé.',
    ],
    [
      'Suppression d’une ligne de tableau quand elle est devient entièrement vide.',
      '(Voir ci dessous)',
      'Si après substitution du motif la cellule plus bas devient vide, toute la ligne du tableau est supprimée.',
    ],
    [
      '',
      '<<Champ 1>>',
      '',
    ],
    [
      'Du texte.',
      '<<Champ 1>>',
      'Cette ligne ne sera pas supprimée, car des cellules contiennent du texte.',
    ]
  ]
  const table = body.appendTable(cells);

  /* styling: */
  /* default style for all paragraphs */
  const baseStyle = {
    [DocumentApp.Attribute.SPACING_AFTER]: 4,
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#FFFFFF',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#000000',
    [DocumentApp.Attribute.BOLD]: false,
  };
  for (const p of body.getParagraphs()) {
    p.setAttributes(baseStyle);
  }
  /* word in bold: "<<Champ 1>>" in row 4 */
  table.getRow(3).getCell(1).editAsText().setBold(9, 19, true);

  /* table header */
  const headStyle = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#4682B4', /* SteelBlue */
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF',
    [DocumentApp.Attribute.BOLD]: true,
  };
  const th = table.getRow(0);
  for (let i=0; i<3; i++) {
    let td = th.getCell(i).setAttributes(headStyle);
    for (let k=0; k<td.getNumChildren(); k++) {
      let e = td.getChild(k); 
      if (e.getType() == DocumentApp.ElementType.PARAGRAPH) {
        e.setAttributes(headStyle);
      }
    }
  }

  doc.saveAndClose();

  return doc;
}



