import docx
from docx.enum.text import WD_COLOR_INDEX

def polish_document(input_path, output_path):
    doc = docx.Document(input_path)
    
    # 定义替换字典：{原文片段: 润色后片段}
    replacements = {
        "face major challenges in integrating": "confront substantial hurdles in the integration of",
        "as a pathway to easing local energy shortages": "as a strategic approach to mitigating regional energy deficits",
        "The entropy weight method (EWM) is used to optimize": "The Entropy Weight Method (EWM) is employed to optimize",
        "identify strong natural seasonal complementarity": "reveal pronounced natural seasonal synergies",
        "provides a quantitative basis": "establishes a rigorous quantitative framework",
        "demonstrates that wind-wave-tidal-solar complementarity can improve energy security and resilience": "illustrates that the synergy among wind, wave, tidal, and solar resources significantly bolsters energy security and resilience",
        "increasingly unable to support sustainable development": "becoming increasingly inadequate for sustaining long-term development",
        "fundamentally limits the reliability of single-resource power systems": "constitutes a fundamental constraint on the reliability of single-source power systems",
        "In response to these challenges, multi-energy complementary systems (MECSs) have been proposed": "To address these complexities, multi-energy complementary systems (MECSs) have emerged as a promising solution",
        "shifting from simply combining multiple energy resources to optimizing the configuration": "transitioning from the rudimentary combination of energy resources toward the sophisticated optimization of system configurations",
        "achieves a better balance between energy utilization and output stability": "attains a more robust equilibrium between energy utilization efficiency and output stability",
        "tends to overconcentrate the portfolio in a single resource": "exhibits a propensity for excessive portfolio concentration in a single resource",
        "provides a more balanced allocation": "facilitates a more equitable and diversified allocation",
        "The reduced seasonal amplitude of the aggregated output suggests a lower reliance": "The attenuated seasonal amplitude of the integrated output implies a diminished dependency",
        "This four-source synergy substantially reduces seasonal fluctuations": "This four-source synergy markedly dampens seasonal fluctuations",
        "offers a transferable and quantitatively grounded framework": "provides a scalable and analytically rigorous framework"
    }

    for paragraph in doc.paragraphs:
        for old_text, new_text in replacements.items():
            if old_text in paragraph.text:
                # 遍历段落中的 runs
                # 注意：如果 old_text 跨越了多个 run，这种简单的替换会失效
                # 但对于大多数学术短语，它们通常在同一个 run 中或我们可以整体替换
                # 为了保留格式，我们尝试在 run 级别操作，或者如果段落结构简单，直接操作整个段落
                
                # 简单起见，如果段落包含该片段，我们直接在 text 中替换
                # 并对整个段落的修改部分进行高亮（这里采用一种折中方案：
                # 如果找到匹配，我们重新构建段落以应用高亮）
                
                original_text = paragraph.text
                if old_text in original_text:
                    # 记录原始格式（简单处理）
                    inline_ref = [] # 记录是否有引文等特殊格式可能比较复杂
                    
                    # 替换文本
                    updated_text = original_text.replace(old_text, new_text)
                    
                    # 清除原段落并重新添加（这会丢失加粗/斜体，但由于我们主要润色文本，这在初稿润色中通常可接受）
                    # 更好的做法是遍历 runs，但那非常复杂
                    
                    # 尝试保留 run 级别的替换以维持格式
                    for run in paragraph.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        elif any(word in run.text for word in old_text.split()):
                            # 处理跨 run 的情况：如果该 run 包含 old_text 的一部分
                            # 这种情况下，简单的 run 级别替换可能不准确
                            pass 

    doc.save(output_path)
    print(f"Document polished and saved to {output_path}")

if __name__ == "__main__":
    polish_document('Manuscript20260412CAY.docx', 'Manuscript20260412CAY.docx')
