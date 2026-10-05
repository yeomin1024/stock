// VERSION: v1.0.0 — 2026-10-05 — 컴포지션 등록
// Molbbang: 본편 (1920×1080, 30fps, 길이 = SRT 마지막 자막 끝 + 고지 3초 + 여유 1초)
// Scenes/ 폴더: 장면별 미리보기 (본편 타임라인을 해당 장면 시작으로 당겨서 그대로 보여줌)
import React from 'react';
import {Composition, Folder, Sequence} from 'remotion';
import {SCENES, FPS, HEIGHT, TOTAL_FRAMES, WIDTH, sceneRange} from './data/timeline';
import {CHANNEL_NAME_DEFAULT} from './data/facts';
import {Video, VideoProps} from './Video';

type ScenePreviewProps = VideoProps & {readonly sceneId: string};

const ScenePreview: React.FC<ScenePreviewProps> = ({sceneId, channelName}) => (
	<Sequence from={-sceneRange(sceneId).start}>
		<Video channelName={channelName} />
	</Sequence>
);

export const RemotionRoot: React.FC = () => (
	<>
		<Composition
			id="Molbbang"
			component={Video}
			durationInFrames={TOTAL_FRAMES}
			fps={FPS}
			width={WIDTH}
			height={HEIGHT}
			defaultProps={{channelName: CHANNEL_NAME_DEFAULT}}
		/>
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
						defaultProps={{sceneId: s.id, channelName: CHANNEL_NAME_DEFAULT}}
					/>
				);
			})}
		</Folder>
	</>
);
