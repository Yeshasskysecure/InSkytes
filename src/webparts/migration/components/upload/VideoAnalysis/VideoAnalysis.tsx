import * as React from "react";
import { useState, useEffect, useRef } from "react";
import styles from './VideoAnalysis.module.scss';
import { AzureOpenAIService } from '../../../services/AzureOpenAIService';
import { buildBuDepartmentValue } from '../../businessUnit/BusinessUnitHierarchyUtils';
import { fetchAllTaxonomyOptions, ITaxonomyFieldOptions, matchTaxonomyTerm } from '../../../services/TaxonomyService';
import { fetchDocumentTypeTerms, Term } from '../../../utils/termStore';

export interface VideoAnalysisProps {
  file: File;
  context?: any;
  onClose?: () => void;
  onAnalysisComplete?: (data: any) => void;
}

interface TranscriptEntry {
  time: number;
  text: string;
}

const cleanTextField = (value: string, maxLength?: number): string => {
  if (!value) return '';

  let cleaned = value
    .replace(/[$%@*+?!#&^~`|\\<>]/g, '')
    .replace(/\.(pdf|docx|mp4|mp3|wav|mov|avi|mkv|webm|m4a|aac|flac|ogg)$/gi, '')
    .replace(/[\s_-]?v\d+(\.\d+)*[\s_-]?/gi, ' ')
    .replace(/[\s_-]?(final|draft|copy|revised|updated|new|old|backup|temp)[\s_-]?/gi, ' ')
    .replace(/[\s_-]?\d{4}[-*]?\d{2}[-*]?\d{2}[\s_-]?/gi, ' ')
    .replace(/[\s_-]#?\d{2,6}[\s_-]?/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();

  if (maxLength && cleaned.length > maxLength) {
    cleaned = cleaned.substring(0, maxLength).trim();
  }

  return cleaned;
};

const buildClientTermsPromptSection = (taxonomyOptions?: ITaxonomyFieldOptions | null): string => {
  const maxClientTermsPromptChars = 24000;
  const clientTerms = (taxonomyOptions?.client || [])
    .map((term) => (term?.label || '').trim())
    .filter((label, index, array) => !!label && array.indexOf(label) === index)
    .sort((left, right) => left.localeCompare(right));

  const lines: string[] = [];
  let totalLength = 0;

  for (let index = 0; index < clientTerms.length; index++) {
    const line = `- "${clientTerms[index]}"`;
    if (totalLength + line.length + 1 > maxClientTermsPromptChars) {
      break;
    }

    lines.push(line);
    totalLength += line.length + 1;
  }

  return lines.join('\n');
};

export const VideoAnalysis: React.FC<VideoAnalysisProps> = ({ file, context, onClose, onAnalysisComplete }) => {
  const isVideo = file.type.startsWith("video");
  const mediaRef = useRef<HTMLVideoElement & HTMLAudioElement>(null);

  const [isProcessing, setIsProcessing] = useState(true);
  const [processingStep, setProcessingStep] = useState(0);
  const [description, setDescription] = useState("");
  const [transcript, setTranscript] = useState<TranscriptEntry[]>([]);
  const [extractedMetadata, setExtractedMetadata] = useState<any>(null);
  const [currentTime, setCurrentTime] = useState(0);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [editValue, setEditValue] = useState("");
  const [taxonomyOptions, setTaxonomyOptions] = useState<ITaxonomyFieldOptions | null>(null);
  const [transcriptionStatus, setTranscriptionStatus] = useState<string>("");
  const [activeTab, setActiveTab] = useState<'abstract' | 'transcript'>('transcript');
  
  const openAIService = useRef(new AzureOpenAIService(undefined, context));
  const transcriptRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let stepTimer: any;
    const processMedia = async () => {
      setProcessingStep(1); 
      try {
        stepTimer = setTimeout(() => setProcessingStep(2), 500);

        const { transcript: finalTranscript, transcriptText } = await openAIService.current.transcribeMedia(
          file,
          (msg) => setTranscriptionStatus(msg)
        );

        setTranscript(finalTranscript || []);
        setIsProcessing(false);
        setActiveTab('abstract');
        setProcessingStep(3);

        let docTypeTermsSection = '';
        let clientTermsSection = '';
        const flatMediaDocTypeTerms: { name: string; description: string }[] = [];
        try {
          if (context?.spHttpClient && context?.pageContext?.web?.absoluteUrl) {
            const mediaDocTypeTerms = await fetchDocumentTypeTerms(
              context.spHttpClient,
              context.pageContext.web.absoluteUrl
            );
            const flattenMediaTerms = (terms: Term[]) => {
              terms.forEach((t) => {
                flatMediaDocTypeTerms.push({
                  name: t.name,
                  description: t.description || ''
                });
                if (t.children?.length) flattenMediaTerms(t.children);
              });
            };
            flattenMediaTerms(mediaDocTypeTerms);
            docTypeTermsSection = flatMediaDocTypeTerms
              .map((t) => (
                t.description
                  ? `- "${t.name}": ${t.description}`
                  : `- "${t.name}"`
              ))
              .join('\n');
          }
        } catch (e) {
          console.warn('Failed to fetch doc type terms for media analysis prompt:', e);
        }
        try {
          if (context) {
            const options = await fetchAllTaxonomyOptions(context);
            setTaxonomyOptions(options);
            clientTermsSection = buildClientTermsPromptSection(options);
          }
        } catch (e) {
          console.warn('Failed to fetch client terms for media analysis prompt:', e);
        }

        const analysisResult = await openAIService.current.generateAbstractFromTranscript(
          file.name,
          transcriptText,
          docTypeTermsSection,
          clientTermsSection
        );

        // Auto-play media when analysis is done
        if (mediaRef.current) {
          mediaRef.current.play().catch(e => console.log('Auto-play blocked:', e));
        }

        const cleanedTitle = cleanTextField(
          (analysisResult.title || file.name)
            .replace(/\.(mp4|mp3|wav|mov|avi|mkv|webm|m4a|aac|flac|ogg)$/gi, '')
            .replace(/[_-]/g, ' ')
            .trim(),
          255
        );

        // Final extracted metadata for form
        const finalMetadata = {
          ...analysisResult,
          title: cleanedTitle,
          description: analysisResult.abstract || '',
          bu: analysisResult.businessUnit || '',
          department: analysisResult.department || '',
          documentType: analysisResult.documentType || ''
        };

        setDescription(finalMetadata.description);
        setExtractedMetadata(finalMetadata);
        setProcessingStep(4);
      } catch (error) {
        console.error("AI Analysis failed:", error);
        setProcessingStep(-1);
        setIsProcessing(false);
      }
    };

    processMedia();
    const allTaxonomy = async () => {
      try {
        const options = await fetchAllTaxonomyOptions(context);
        setTaxonomyOptions(options);
      } catch (err) {
        console.warn('Failed to fetch taxonomy for media analysis:', err);
      }
    };
    allTaxonomy();

    return () => {
      if (stepTimer) clearTimeout(stepTimer);
    };
  }, [file]);

  useEffect(() => {
    const media = mediaRef.current;
    if (!media) return;

    const handleTimeUpdate = () => {
      setCurrentTime(media.currentTime);
    };

    media.addEventListener('timeupdate', handleTimeUpdate);
    return () => media.removeEventListener('timeupdate', handleTimeUpdate);
  }, []);

  const activeIndex = transcript.findIndex((entry, index) => {
    const nextEntry = transcript[index + 1];
    return currentTime >= entry.time && (!nextEntry || currentTime < nextEntry.time);
  });

  useEffect(() => {
    if (!transcriptRef.current || activeIndex === -1) return;
    
    // Slight delay to ensure the row has rendered its active class if needed
    setTimeout(() => {
      if (transcriptRef.current) {
        const activeRow = transcriptRef.current.querySelector(`.${styles.activeTranscriptRow}`);
        if (activeRow) {
          activeRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }
    }, 50);
  }, [activeIndex, activeTab]);


  const seekTo = (time: number) => {
    if (mediaRef.current) {
      mediaRef.current.currentTime = time;
      mediaRef.current.play().catch(() => {});
    }
  };

  const startEditing = (index: number, text: string) => {
    setEditingIndex(index);
    setEditValue(text);
  };

  const saveEdit = (index: number) => {
    const newTranscript = [...transcript];
    newTranscript[index].text = editValue;
    setTranscript(newTranscript);
    setEditingIndex(null);
  };

  const formatTime = (seconds: number) => {
    const m = Math.floor(seconds / 60);
    const s = Math.floor(seconds % 60);
    return `${m}:${s.toString().padStart(2, '0')}`;
  };

  const formatBytes = (bytes: number, decimals = 2) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  };

  return (
    <div className={styles.modalContainer}>
      <div className={styles.header}>
        <div className={styles.headerLeft}>
          <span className={styles.headerIcon}>🎬</span>
          <div>
            <h2 className={styles.headerTitle}>Media Analysis</h2>
            <p className={styles.headerSubtext}>{isVideo ? 'Video' : 'Audio'} - {formatBytes(file.size)}</p>
          </div>
        </div>
        {onClose && (
          <button className={styles.closeBtn} onClick={onClose} aria-label="Close">✕</button>
        )}
      </div>

      <div className={styles.body}>
        <div className={styles.leftColumn}>
            <div className={styles.mediaHeaderInfo}><strong>File: </strong> {file.name}</div>
            <div className={styles.mediaHeaderInfo}><strong>Size: </strong> {formatBytes(file.size)}</div>
            
            <div className={styles.playerContainer}>
              {isVideo ? (
                <video ref={mediaRef} controls className={styles.videoPlayer}>
                  <source src={URL.createObjectURL(file)} />
                </video>
              ) : (
                <audio ref={mediaRef} controls className={styles.videoPlayer}>
                  <source src={URL.createObjectURL(file)} />
                </audio>
              )}
            </div>
        </div>

        <div className={styles.rightColumn}>
          {isProcessing ? (
            <div className={styles.processingBlock}>
               <div className={styles.spinnerWrapper}>
                 <div className={styles.spinner}></div>
               </div>
               <h3 className={styles.processingHeader}>Analysis in Progress</h3>
               <div className={styles.processingStepsCard}>
                 <div className={processingStep >= 1 ? styles.stepDone : styles.stepPending}>
                   <div className={styles.stepIndicator}></div>
                   <div className={styles.stepText}>Media uploaded successfully</div>
                 </div>
                 <div className={processingStep >= 2 ? styles.stepDone : styles.stepPending}>
                   <div className={styles.stepIndicator}></div>
                   <div className={styles.stepText}>
                     Generating transcript 
                     {transcriptionStatus && <span className={styles.subStatus}>{transcriptionStatus}</span>}
                   </div>
                 </div>
                 <div className={processingStep >= 3 ? styles.stepDone : styles.stepPending}>
                   <div className={styles.stepIndicator}></div>
                   <div className={styles.stepText}>Drafting detailed summary</div>
                 </div>
               </div>
            </div>
          ) : (
            <div className={styles.contentWrap}>
              <div className={styles['tabsContainer']}>
                <div 
                  className={`${styles['tab']} ${activeTab === 'transcript' ? styles['activeTab'] : ''}`}
                  onClick={() => setActiveTab('transcript')}
                >
                  Transcription
                </div>
                <div 
                  className={`${styles['tab']} ${activeTab === 'abstract' ? styles['activeTab'] : ''}`}
                  onClick={() => setActiveTab('abstract')}
                >
                  Description
                </div>
              </div>

              {activeTab === 'abstract' ? (
                <div className={styles.sectionBlock}>
                  <div className={styles.abstractBox}>
                    {processingStep === 3 ? (
                      <div className={styles.generatingAbstractMessage}>
                        <div className={styles.spinnerSmall}></div>
                        <span>Structuring insights and drafting summary...</span>
                      </div>
                    ) : (
                      <div className={styles['corporateAbstract']}>
                        {description}
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div className={styles.sectionBlock}>
                  <div className={styles.transcriptWrap} ref={transcriptRef}>
                    {transcript.map((entry, index) => {
                      const nextEntry = transcript[index + 1];
                      const isActive = currentTime >= entry.time && (!nextEntry || currentTime < nextEntry.time);
                      
                      return (
                        <div 
                          key={index} 
                          className={`${styles.transcriptRow} ${isActive ? styles.activeTranscriptRow : ''}`}
                          onClick={() => seekTo(entry.time)}
                        >
                          <span className={styles.transcriptTime}>[{formatTime(entry.time)}]</span>
                          {editingIndex === index ? (
                            <input 
                              autoFocus
                              className={styles.editInput}
                              value={editValue}
                              onChange={(e) => setEditValue(e.target.value)}
                              onBlur={() => saveEdit(index)}
                              onKeyDown={(e) => { if (e.key === 'Enter') saveEdit(index); if (e.key === 'Escape') setEditingIndex(null); }}
                              onClick={(e) => e.stopPropagation()}
                            />
                          ) : (
                            <>
                              <span className={styles.transcriptTextRow}>{entry.text}</span>
                              <div className={styles.transcriptActions}>
                                <button className={styles.editBtn} onClick={(e) => { e.stopPropagation(); startEditing(index, entry.text); }}>Edit</button>
                              </div>
                            </>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <div className={styles.footer}>
        <button className={styles.btnCancel} onClick={onClose}>Cancel</button>
        <button 
          className={styles.btnUpload} 
          disabled={isProcessing}
          onClick={() => {
            if (onAnalysisComplete && extractedMetadata) {
              onAnalysisComplete(extractedMetadata);
            }
          }}
        >
          Continue to Metadata
        </button>
      </div>
    </div>
  );
};

export default VideoAnalysis;

