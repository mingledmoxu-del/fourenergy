import docx
from docx.enum.text import WD_COLOR_INDEX

def deep_polish(input_path, output_path):
    doc = docx.Document(input_path)
    
    # 深度润色映射字典，涵盖摘要、引言、讨论及全文高频弱词
    polish_map = {
        # 摘要 & 引言核心句式优化
        "face major challenges in integrating": "confront formidable technical hurdles in the seamless integration of",
        "as a pathway to easing local energy shortages": "as a strategic paradigm to alleviate regional energy deficits",
        "The entropy weight method (EWM) is used to optimize": "The Entropy Weight Method (EWM) is meticulously employed to systematically optimize",
        "identify strong natural seasonal complementarity": "unveil pronounced and inherent seasonal synergies",
        "provides a quantitative basis": "establishes a robust and rigorous quantitative framework",
        "demonstrates that wind-wave-tidal-solar complementarity can improve energy security and resilience": "substantiates that the synergistic orchestration of wind, wave, tidal, and solar resources significantly bolsters energy security and systemic resilience",
        "increasingly unable to support sustainable development": "becoming increasingly insufficient for sustaining long-term developmental trajectories",
        "placed renewed emphasis on the use of": "precipitated a renewed strategic emphasis on the utilization of",
        "fundamentally limits the reliability of": "constitutes a fundamental constraint on the operational reliability of",
        "In response to these challenges, multi-energy complementary systems (MECSs) have been proposed": "To address these multifaceted complexities, multi-energy complementary systems (MECSs) have emerged as a pivotal solution",
        "shifting from simply combining multiple energy resources to optimizing the configuration": "transitioning from the rudimentary aggregation of energy resources toward the sophisticated multi-objective optimization of system architectures",
        "effectively reduced dependence on fossil fuels": "markedly mitigated the reliance on conventional fossil-fuel-based power generation",
        
        # 结果与讨论深度润色
        "achieves a better balance between": "attains a more optimal equilibrium between",
        "tends to overconcentrate the portfolio": "demonstrates a propensity for excessive and suboptimal portfolio concentration",
        "The key difference between EWM and MSD is not whether variability can be minimized": "The fundamental distinction between EWM and MSD lies not merely in the minimization of variability",
        "spatial analysis further identifies": "comprehensive spatial analysis further elucidates",
        "highlights the need for flexible loads": "underscores the imperative for integrating flexible demand-side resources",
        "The reduced seasonal amplitude of the aggregated output suggests": "The attenuated seasonal amplitude of the integrated output implies",
        "This result shows that the complementarity among natural resources can reduce dependence": "This finding indicates that the natural synergies among disparate resources can partially offset the requirement",
        "improving overall system operating efficiency": "enhancing the holistic operational efficiency of the system",
        "optimization aimed solely at reducing volatility is not suitable": "optimization frameworks exclusively focused on volatility reduction are inadequate",
        
        # 词汇与表达升级 (Global Weak Words)
        "shows that": "indicates that",
        "find out": "elucidate",
        "deal with": "address",
        "improve": "enhance",
        "big": "substantial",
        "small": "marginal",
        "good": "favorable",
        "bad": "detrimental",
        "important": "pivotal",
        "use": "utilize",
        "change": "variation",
        "get": "obtain",
        "result in": "precipitate",
        "because of": "owing to",
        "enough": "sufficient",
        "about": "approximately",
        "very": "considerably",
        "many": "numerous",
        
        # 句式连贯性与深度
        "In summary, the renewable resources of": "Synthetically, the renewable resource profile of",
        "Overall, the findings confirm": "Collectively, the empirical findings validate",
        "Future research should incorporate": "Prospective inquiries should integrate"
    }

    modified_count = 0
    for paragraph in doc.paragraphs:
        # 对每一个段落进行多轮替换，以确保长句和短语都能被覆盖
        # 我们采用逆序长度排序，防止短词替换破坏长句模式
        sorted_keys = sorted(polish_map.keys(), key=len, reverse=True)
        
        original_text = paragraph.text
        for old_text in sorted_keys:
            new_text = polish_map[old_text]
            if old_text in paragraph.text:
                # 为了保留尽可能多的格式，我们遍历 runs
                for run in paragraph.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        modified_count += 1
                
                # 如果 old_text 跨越了 runs，paragraph.text 会显示存在，但单个 run.text 可能没有
                # 这种情况下我们需要更复杂的处理。为确保覆盖，我们检查段落文本。
                if old_text in paragraph.text and new_text not in paragraph.text:
                    # 这是一个简单的跨 run 修复：如果还没被替换成功，直接重写段落并标黄
                    # 注意：这会丢失段落内的细微格式（如某个单词加粗），但保证了润色效果和高亮
                    paragraph.text = paragraph.text.replace(old_text, new_text)
                    for run in paragraph.runs:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    modified_count += 1

    doc.save(output_path)
    print(f"Deep polish completed. {modified_count} occurrences addressed and highlighted.")

if __name__ == "__main__":
    deep_polish('Manuscript20260412CAY.docx', 'Manuscript20260412CAY.docx')
