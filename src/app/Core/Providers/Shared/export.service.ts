import { Injectable } from '@angular/core';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';

@Injectable({
    providedIn: 'root'
})
export class PdfExportService {

    async exportToPDF(
        elementId: string,
        fileName: string = 'Report',
        orientation: 'p' | 'l' = 'l'
    ) {
        const element = document.getElementById(elementId);

        if (!element) {
            console.error('Element not found:', elementId);
            return;
        }

        // 🔥 Clone element (avoid UI breaking)
        const clonedElement = element.cloneNode(true) as HTMLElement;

        this.applyPdfStyles(clonedElement);

        document.body.appendChild(clonedElement);

        const canvas = await html2canvas(clonedElement, {
            scale: 2,
            useCORS: true,
            scrollX: 0,
            scrollY: 0
        });

        document.body.removeChild(clonedElement);

        const imgData = canvas.toDataURL('image/png');

        const pdf = new jsPDF(orientation, 'mm', 'a4');

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        const margin = 5;
        const imgWidth = pageWidth - margin * 2;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        let heightLeft = imgHeight;
        let position = margin;

        // First Page
        pdf.addImage(imgData, 'PNG', margin, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;

        // Multi-page
        while (heightLeft > 0) {
            position = heightLeft - imgHeight;
            pdf.addPage();
            pdf.addImage(imgData, 'PNG', margin, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
        }

        pdf.save(`${fileName}.pdf`);
    }

    private applyPdfStyles(element: HTMLElement) {

        // Move off-screen
        element.style.position = 'fixed';
        element.style.left = '-9999px';
        element.style.top = '0';
        element.style.width = '1400px'; // 🔥 keep wide layout

        // Remove scroll containers
        element.querySelectorAll(
            '.scorecard-multi-header, .scorecard-triple-header, .scorecard-double-header, .scorecard-single-header, .scorecard-aging-header'
        ).forEach((el: any) => {
            el.style.height = 'auto';
            el.style.overflow = 'visible';
            el.style.maxWidth = 'none';
        });

        // Fix table width (IMPORTANT)
        element.querySelectorAll('.table').forEach((table: any) => {
            table.style.width = '100%';              // 🔥 no scroll
            table.style.maxWidth = '100%';
            table.style.tableLayout = 'auto';        // keep natural column width
        });

        // Remove sticky (but KEEP look)
        element.querySelectorAll('th, td').forEach((cell: any) => {
            cell.style.position = 'static';
            cell.style.whiteSpace = 'nowrap'; // keep same UI look
        });

        // Remove sticky headers/footers
        element.querySelectorAll('thead, .sticky-footer, .sticky-bottom-row')
            .forEach((el: any) => {
                el.style.position = 'static';
            });

        // Prevent clipping
        element.querySelectorAll('*').forEach((el: any) => {
            el.style.overflow = 'visible';
        });
    }
}