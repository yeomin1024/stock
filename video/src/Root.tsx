// VERSION: v2.4.0 — 2026-10-08 — 컴포지션 등록 (주석만 갱신: 길이 계산)
// Molbbang: 본편 (1920×1080, 30fps, 길이 = SRT 마지막 자막 끝 + 여유 1초 — v2.4.0 부터 SRT 에 고지 카드 3.5초 공백 포함)
// Scenes/ 폴더: 장면별 미리보기 (본편 타임라인을 해당 장면 시작으로 당겨서 그대로 보여줌)
import React from 'react';
import {Composition, Folder, Sequence} from 'remotion';
import {SCENES, FPS, HEIGHT, TOTAL_FRAMES, WIDTH, sceneRange} from './data/timeline';
import {Video} from './Video';

const ScenePreview: React.FC<{readonly sceneId: string}> = ({sceneId}) => (
	<Sequence from={-sceneRange(sceneId).start}>
		<Video />
	</Sequence>
);

export const RemotionRoot: React.FC = () => (
	<>
		<Composition id="Molbbang" component={Video} durationInFrames={TOTAL_FRAMES} fps={FPS} width={WIDTH} height={HEIGHT} />
		<Folder name="Scenes">
			{SCENES.map((s) => {
				const {start, end} = sceneRange(s.id);
				return (
					<Composition
						key={s.id}
						id={s.id}
						component={ScenePreview}
						durationInFrames={end - start}
						fps={FPS}
						width={WIDTH}
						height={HEIGHT}
						defaultProps={{sceneId: s.id}}
					/>
				);
			})}
		</Folder>
	</>
);
