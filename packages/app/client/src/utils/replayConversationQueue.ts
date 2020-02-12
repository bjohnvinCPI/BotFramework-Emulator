//
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
// Microsoft Bot Framework: http://botframework.com
//
// Bot Framework Emulator Github:
// https://github.com/Microsoft/BotFramwork-Emulator
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
import { EventEmitter } from 'events';

import { Activity } from 'botframework-schema';
import groupBy from 'lodash.groupby';
import { Store } from 'redux';
import moment from 'moment';
import { call } from 'redux-saga/effects';

import { RootState } from '../state';

export const INCOMING_ACTIVITY_EVENT = 'INCOMING_ACTIVITY';
export const COMPLETED_ACTIVITY_EVENT = 'COMPLETED_ACTIVITY';

interface ActivityIdMap<T> {
  [x: string]: T;
}

interface WebchatStore {
  dispatch: Function;
}

class ActivityEmitter extends EventEmitter {}

async function asyncForEach(array, callback) {
  for (let index = 0; index < array.length; index++) {
    await callback(array[index], index, array);
  }
}

export class ConversationQueue {
  private originalActivityQueue: ActivityIdMap<Activity[]>;
  private replayedActivityQueue: ActivityIdMap<Activity[]>;
  private activitiesToBeReplayed: Activity[] = [];
  private conversationId: string = '';
  private addActivityEmitter: EventEmitter;
  private webchatStore: WebchatStore;
  private currentActivityListener: EventEmitter;

  constructor(conversationId: string, originalActivities: Activity[]) {
    this.conversationId = conversationId;
    this.addActivityEmitter = new EventEmitter();
    this.originalActivityQueue = {};
    this.replayedActivityQueue = {};

    this.onIncomingEvent = this.onIncomingEvent.bind(this);
    this.handleActivityEmittedEvents = this.handleActivityEmittedEvents.bind(this);
    this.buildBotActivityQueue = this.buildBotActivityQueue.bind(this);
    this.processActivities = this.processActivities.bind(this);
    this.processActivity = this.processActivity.bind(this);
    this.updateStore = this.updateStore.bind(this);
    this.onEventCompleted = this.onEventCompleted.bind(this);

    this.buildActivitiesToBeReplayed(originalActivities);
    this.buildBotActivityQueue(originalActivities);
    this.handleActivityEmittedEvents();
  }

  public updateStore(webchatStore: Store<RootState>) {
    this.webchatStore = webchatStore;
  }

  private buildActivitiesToBeReplayed(originalActivities: Activity[]) {
    const userActivities: Activity[] = originalActivities.filter(
      activity => activity.from.role === 'user' && activity.type === 'message'
    );

    this.activitiesToBeReplayed = userActivities.map((activity: Activity) => {
      return {
        ...activity,
        conversation: {
          ...activity.conversation,
          id: this.conversationId,
        },
        channelData: {
          ...activity.channelData,
          originalActivityId: activity.id,
        },
      };
    });
  }

  private buildBotActivityQueue(originalActivities: Activity[]) {
    const botActivities: Activity[] = originalActivities.filter(activity => activity.from.role !== 'user');
    this.originalActivityQueue = groupBy(botActivities, 'replyToId');
  }

  private onEventCompleted(evts) {
    this.currentActivityListener.emit('completed', evts);
  }

  private onIncomingEvent(currentActivity) {
    if (currentActivity.channelData) {
      const originalActivityId = currentActivity.channelData.originalActivityId;
      if (originalActivityId) {
        if (!this.replayedActivityQueue[originalActivityId]) {
          this.replayedActivityQueue[originalActivityId] = [];
        }
        this.replayedActivityQueue[originalActivityId].push(currentActivity);
      }
    }
  }

  private handleActivityEmittedEvents() {
    this.addActivityEmitter.on(INCOMING_ACTIVITY_EVENT, this.onIncomingEvent);
    this.addActivityEmitter.on(COMPLETED_ACTIVITY_EVENT, this.onEventCompleted);
  }

  public get activityEmitter(): EventEmitter {
    return this.addActivityEmitter;
  }

  private dataURLtoFile(dataurl, filename) {
    var arr = dataurl.split(','),
      mime = arr[0].match(/:(.*?);/)[1],
      bstr = atob(arr[1]),
      n = bstr.length,
      u8arr = new Uint8Array(n);

    while (n--) {
      u8arr[n] = bstr.charCodeAt(n);
    }

    return new File([u8arr], filename, { type: mime });
  }

  private async processActivity(activity: Activity) {
    return new Promise((resolve, reject) => {
      try {
        if (activity.attachments && activity.attachments.length >= 1) {
          const mutatedAttachments = activity.attachments.map(attachment => {
            const fileFormat: File = this.dataURLtoFile(attachment.contentUrl, attachment.name);
            return {
              ...attachment,
              contentUrl: window.URL.createObjectURL(fileFormat),
            };
          });
          activity.attachments = mutatedAttachments;
        }

        this.webchatStore.dispatch({
          type: 'DIRECT_LINE/POST_ACTIVITY',
          payload: {
            activity,
          },
          meta: {
            method: 'keyboard',
          },
        });
        this.currentActivityListener = new ActivityEmitter();
        this.currentActivityListener.once('completed', async result => {
          await new Promise(resolve => {
            setTimeout(() => {
              resolve();
            }, 1000);
          });
          resolve('success');
        });
      } catch (ex) {
        console.log(ex);
        reject(ex);
      }
    });
  }

  public processActivities() {
    return new Promise((resolve, reject) => {
      try {
        asyncForEach(this.activitiesToBeReplayed, async (activity: Activity, index: number) => {
          try {
            const resp = await this.processActivity(activity);
            console.log(resp);
          } catch (ex) {
            console.log(ex);
          }
          if (index === this.activitiesToBeReplayed.length - 1) {
            resolve();
          }
        });
      } catch (ex) {
        console.log(ex);
        reject();
      }
    });
  }
}
