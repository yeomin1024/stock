// VERSION: v1.0.0 — 2026-10-05 — 채널명 변수 (컴포지션 prop → 컨텍스트). 기본값은 대본 표기 "[채널명]".
import React, {createContext, useContext} from 'react';
import {CHANNEL_NAME_DEFAULT} from './facts';

const ChannelCtx = createContext<string>(CHANNEL_NAME_DEFAULT);

export const ChannelProvider: React.FC<{readonly name: string; readonly children: React.ReactNode}> = ({name, children}) => (
	<ChannelCtx.Provider value={name.trim() === '' ? CHANNEL_NAME_DEFAULT : name}>{children}</ChannelCtx.Provider>
);

export const useChannelName = (): string => useContext(ChannelCtx);
