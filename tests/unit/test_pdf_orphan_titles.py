# Copyright (C) 2026 Aguirre MAURIN
#
# Ce programme est un logiciel libre : vous pouvez le redistribuer et/ou le modifier
# selon les termes de la Licence Publique Générale GNU (GPL) telle que publiée par
# la Free Software Foundation, version 3 de la licence, ou (à votre choix) toute version ultérieure.
#
# Ce programme est distribué dans l'espoir qu'il sera utile, mais SANS AUCUNE GARANTIE ;
# sans même la garantie implicite de QUALITÉ MARCHANDE ou D'ADÉQUATION À UN USAGE PARTICULIER.
# Voir la Licence Publique Générale GNU pour plus de détails.
#
# CONDITIONS SUPPLÉMENTAIRES D'ATTRIBUTION (SECTION 7(b) DE LA GPL v3) :
# Conformément à la section 7(b) de la GNU GPL v3, vous devez expressément conserver
# intactes et lisibles toutes les mentions d'auteur, notices de copyright et la présente
# clause dans chaque fichier source ou interface utilisateur redistribué. Toute version modifiée
# doit clairement indiquer qu'elle a été altérée et ne doit en aucun cas supprimer le nom
# de l'auteur original (Aguirre MAURIN).

#
import pytest
from reportlab.platypus import Paragraph, Spacer, Image as RLImage, KeepTogether
from core.common.pdf_report_builder import PDFReportBuilder
from reportlab.lib.styles import getSampleStyleSheet
import tempfile
from pathlib import Path

def test_no_orphan_titles_in_keeptogether():
    """
    Test guard pour éviter qu'un bloc contenant uniquement [Titre, Spacer]
    ne soit enfermé dans un KeepTogether, ce qui annule le keepWithNext
    et provoque des titres orphelins en bas de page.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "test.pdf"
        builder = PDFReportBuilder(pdf_path, header_title="Test")
        
        styles = getSampleStyleSheet()
        title_para = Paragraph("Titre test", styles["Heading1"])
        spacer = Spacer(1, 10)
        from unittest.mock import MagicMock
        img = MagicMock(spec=RLImage)
        
        block = [title_para, spacer, img]
        
        # Test leading chunk length logic
        # Should return 2 (title + spacer) so the image is not included
        chunk_len = builder._leading_title_chunk_len(block)
        assert chunk_len == 2
        
        prefix = block[:chunk_len]
        
        # Le test doit s'assurer que si on appelle _append_with_pending avec keep_together=True,
        # le préfixe (qui n'a QUE le titre et l'espace) ne finit pas bêtement dans un KeepTogether
        # sinon on perd l'effet keepWithNext !
        
        builder._append_with_pending(block, keep_together=True)
        
        for item in builder.story:
            if isinstance(item, KeepTogether):
                # Vérification simplifiée: si item correspond exactement à [Titre, Spacer]
                content = getattr(item, "_content", [])
                if len(content) == 2 and isinstance(content[0], Paragraph) and isinstance(content[1], Spacer):
                    pytest.fail("Un [Titre, Spacer] a été enfermé dans un KeepTogether, ce qui crée un titre orphelin.")