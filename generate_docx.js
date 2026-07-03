const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat } = require('docx');

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, color: "1E40AF" },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, color: "1E40AF" },
        paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, color: "1E40AF" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } }
    ]
  },
  numbering: {
    config: [
      { reference: "bullet-list",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbered-list",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]
  },
  sections: [{
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Contenidos Web FiscalPro")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1. Página de Inicio")] }),
      new Paragraph({ children: [new TextRun({ text: "Título: ", bold: true }), new TextRun("Tu tranquilidad fiscal es nuestro negocio")] }),
      new Paragraph({ children: [new TextRun("Subtítulo: Ayudamos a autónomos y PYMES a navegar por la complejidad del sistema tributario español con soluciones personalizadas.")] }),
      new Paragraph({ children: [new TextRun({ text: "Sección Servicios:", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Autónomos: Gestión de altas, IVA, IRPF y asesoramiento continuo en tu actividad diaria.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Empresas: Impuesto de Sociedades, Cuentas Anuales y planificación fiscal estratégica para PYMES.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Auditoría: Revisión de estados contables y defensa ante inspecciones de la AEAT.")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("2. Servicios Detallados")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Gestión Integral para Autónomos")] }),
      new Paragraph({ children: [new TextRun("Nos encargamos de todo el ciclo de vida fiscal de tu actividad. Desde el alta inicial hasta la presentación de impuestos trimestrales y anuales.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Alta en RETA y Hacienda (Modelos 036/037)")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Liquidaciones de IVA (Modelo 303) e IRPF (Modelo 130)")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Libros Registro oficiales")] }),
      
      new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("Asesoría para PYMES y Sociedades")] }),
      new Paragraph({ children: [new TextRun("Optimizamos la estructura fiscal de tu empresa para maximizar el ahorro y asegurar el cumplimiento normativo.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Impuesto de Sociedades (Modelo 200)")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Depósito de Cuentas Anuales")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Planificación fiscal estratégica")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3. Guía: Gastos Deducibles")] }),
      new Paragraph({ children: [new TextRun("Para que un gasto sea deducible en el IRPF y el IVA, debe cumplir tres requisitos fundamentales:")] }),
      new Paragraph({ numbering: { reference: "numbered-list", level: 0 }, children: [new TextRun("Estar vinculado a la actividad económica.")] }),
      new Paragraph({ numbering: { reference: "numbered-list", level: 0 }, children: [new TextRun("Estar justificado mediante factura.")] }),
      new Paragraph({ numbering: { reference: "numbered-list", level: 0 }, children: [new TextRun("Estar registrado en los libros contables.")] }),
      new Paragraph({ children: [new TextRun({ text: "Categorías:", bold: true })] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Consumos de explotación: Compra de materiales.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Arrendamientos: Alquiler de local u oficina.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Servicios profesionales: Gestores, abogados.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Manutención: Comidas de trabajo con límites diarios.")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4. Guía: El IVA")] }),
      new Paragraph({ children: [new TextRun("Impuesto indirecto que recae sobre el consumo. Actúas como recaudador para Hacienda.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("General (21%): Mayoría de productos.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Reducido (10%): Hostelería, alimentos.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Superreducido (4%): Alimentos básicos, libros.")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("5. Alta en el RETA")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Alta en Hacienda: Modelos 036/037.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Alta en Seguridad Social: En los 60 días previos.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Cuota: Cotización por ingresos reales.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Tarifa Plana: 80€/mes el primer año.")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("6. Deducciones Comunidad Valenciana")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Vivienda: Alquiler e instalaciones de autoconsumo eléctrico.")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Salud: Óptica, dentista y salud mental (30%).")] }),
      new Paragraph({ numbering: { reference: "bullet-list", level: 0 }, children: [new TextRun("Deporte: Gimnasios y federaciones.")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7. Aviso Legal")] }),
      new Paragraph({ children: [new TextRun("Titular: [Tu Nombre o Empresa]")] }),
      new Paragraph({ children: [new TextRun("NIF/CIF: [Tu NIF/CIF]")] }),
      new Paragraph({ children: [new TextRun("La información fiscal es meramente informativa.")] })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Contenidos_Web_FiscalPro.docx", buffer);
  console.log("Archivo creado correctamente.");
});
